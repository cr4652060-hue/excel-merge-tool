package com.example.excelmerge.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;
import java.util.regex.Pattern;

@Service
public class ExcelMergeService {
    private static final int HEADER_SCAN_LIMIT = 30;
    private static final int TYPE_SAMPLE_LIMIT = 50;
    private static final int PREVIEW_LIMIT = 500;
    private static final int INVALID_ROW_LIMIT = 30;

    private static final Pattern HEADER_TEXT_PATTERN = Pattern.compile(".*[A-Za-z\\u4e00-\\u9fff].*");
    private static final Pattern SERIAL_HEADER_PATTERN = Pattern.compile("^(序号|序|编号|行号|序列|no|No|NO)$");
    private static final Pattern FIXED_VALUE_HEADER_PATTERN = Pattern.compile(".*(账户类型|账户类别).*");
    private static final Pattern NON_CORE_HEADER_PATTERN = Pattern.compile(".*(备注|说明|填报人|填表人|填报日期|填表日期).*");
    // ✅ 新增：说明行关键词
    private static final Pattern INSTRUCTION_KEYWORDS = Pattern.compile(
            ".*(填写|说明|注意|示例|要求|口径|备注|提示|温馨提示|如实|以下|请按|请填写|填报|填表|规则|校验|检查).*"
    );
    private static final List<String> KEY_FIELD_KEYWORDS = List.of(
            "设备类型及名称",
            "设备类型名称",
            "设备类型",
            "规格型号",
            "设备序列号",
            "设备序号",
            "管理人",
            "使用人"
    );
    private final AtomicReference<TemplateDefinition> templateRef = new AtomicReference<>();
    private final AtomicReference<List<List<String>>> mergedRowsRef = new AtomicReference<>();
    private final List<TemplateRule> templateRules = loadTemplateRules();

    public TemplateInfo analyzeTemplate(MultipartFile file) {
        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = pickBestDataSheet(workbook);

            int headerRow = findHeaderRowByDensity(sheet);
            if (headerRow < 0) {
                headerRow = findFirstNonEmptyRow(sheet);
            }
            if (headerRow < 0) {
                throw new IllegalStateException("未找到表头行，请检查模板内容。");
            }

            Row row = sheet.getRow(headerRow);
            if (row == null) {
                throw new IllegalStateException("表头行为空，请检查模板内容。");
            }

            List<String> headers = new ArrayList<>();
            List<String> normalized = new ArrayList<>();
            List<Integer> columnIndexes = new ArrayList<>();

            DataFormatter fmt = new DataFormatter();
            for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                String name = cell == null ? "" : fmt.formatCellValue(cell).trim();
                if (name.isBlank()) {
                    continue;
                }
                headers.add(name);
                normalized.add(normalizeHeader(name));
                columnIndexes.add(c);
            }

            if (headers.isEmpty()) {
                throw new IllegalStateException("模板表头没有有效列，请检查模板内容。");
            }

            List<ColumnType> types = detectColumnTypes(sheet, headerRow + 1, columnIndexes);
            Set<String> requiredNormalizedHeaders = resolveRequiredHeaders(normalized);
            TemplateDefinition definition = new TemplateDefinition(
                    headers,
                    normalized,
                    types,
                    requiredNormalizedHeaders,
                    headerRow,
                    headerRow + 1
            );
            templateRef.set(definition);
            mergedRowsRef.set(null);

            return new TemplateInfo(headers, headerRow + 1, headerRow + 2, types);
        } catch (IOException e) {
            throw new IllegalStateException("模板解析失败：" + e.getMessage(), e);
        }
    }

    public MergeResult mergeFiles(List<MultipartFile> files) {
        TemplateDefinition template = templateRef.get();
        if (template == null) {
            throw new IllegalStateException("请先上传模板文件，再进行合并。");
        }
        if (files == null || files.isEmpty()) {
            throw new IllegalArgumentException("请至少上传一份支行 Excel。");
        }

        List<List<String>> mergedRows = new ArrayList<>();
        List<MergeIssue> issues = new ArrayList<>();

        for (MultipartFile file : files) {
            if (file.isEmpty()) {
                issues.add(new MergeIssue(file.getOriginalFilename(), null, null, null, "文件为空，已跳过。"));
                continue;
            }
            try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
                Sheet sheet = pickBestDataSheet(workbook);


                int headerRowIndex = findHeaderRowByMatch(sheet, template.normalizedHeaders());
                if (headerRowIndex < 0) {
                    issues.add(new MergeIssue(file.getOriginalFilename(), sheet.getSheetName(), null, null,
                            "未找到匹配模板的表头，已跳过。"));
                    continue;
                }

                ColumnMapping columnMapping = buildColumnMapping(sheet.getRow(headerRowIndex));
                Map<String, Integer> columnMap = columnMapping.columnMap();
                if (!columnMapping.duplicateHeaders().isEmpty()) {
                    for (String duplicate : columnMapping.duplicateHeaders()) {
                        issues.add(new MergeIssue(file.getOriginalFilename(), sheet.getSheetName(), null,
                                resolveHeaderName(template, duplicate), "列重复，已跳过该文件"));
                    }
                    continue;
                }
                Set<String> missingColumns = new HashSet<>();
                for (int i = 0; i < template.normalizedHeaders().size(); i++) {
                    String norm = template.normalizedHeaders().get(i);
                    if (!columnMap.containsKey(norm)) {
                        missingColumns.add(norm);
                        issues.add(new MergeIssue(file.getOriginalFilename(), sheet.getSheetName(), null,
                                template.headers().get(i), "缺少列：" + template.headers().get(i)));
                    }
                }

                DataFormatter fmt = new DataFormatter();
                List<String> coreHeaders = resolveCoreHeaders(template);
                Integer serialColumn = resolveSerialColumnIndex(template, columnMap);
                List<Integer> keyColumns = resolveKeyColumns(template, columnMap);
                int invalidStreak = 0;
                for (int r = headerRowIndex + 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    // ✅ 1) 处理空行
                    if (row == null) {
                        if (shouldStopByInvalidRow(++invalidStreak)) {
                            break;
                        }
                        continue;
                    }

                    // ✅ 2) 跳过筛选隐藏行（只合并“可见行”）
                    if (isHiddenRow(row)) {
                        if (shouldStopByInvalidRow(++invalidStreak)) {
                            break;
                        }
                        continue;
                    }

                    // ✅ 3) 仅符合业务规则的行才算数据行
                    if (!isBusinessDataRow(row, fmt, serialColumn, keyColumns,
                            coreHeaders, template.requiredNormalizedHeaders(), columnMap)) {
                        if (shouldStopByInvalidRow(++invalidStreak)) {
                            break;
                        }
                        continue;
                    }
                    invalidStreak = 0;
                    List<String> values = new ArrayList<>();
                    for (int c = 0; c < template.normalizedHeaders().size(); c++) {
                        String norm = template.normalizedHeaders().get(c);
                        ColumnType expectedType = template.columnTypes().get(c);
                        Integer colIdx = columnMap.get(norm);
                        String value = "";
                        Cell cell = null;
                        if (colIdx != null && row != null) {
                            cell = row.getCell(colIdx);
                            value = cell == null ? "" : fmt.formatCellValue(cell).trim();
                        }
                        values.add(value);

                        if (missingColumns.contains(norm)) {
                            continue;
                        }
                        // =========================
// 按校验等级处理空值
// =========================
                        if (isSerialHeader(norm)) {
                            continue;
                        }
                        if (value.isBlank()) {
                            if (validationLevel == ValidationLevel.STRICT
                                    && template.requiredNormalizedHeaders().contains(norm)) {

                                issues.add(new MergeIssue(
                                        file.getOriginalFilename(),
                                        sheet.getSheetName(),
                                        r + 1,
                                        template.headers().get(c),
                                        "必填项为空"
                                ));
                            }
                            // 不管严格还是宽松，空值都不再做类型校验
                            continue;
                        }

// =========================
// 只有“有值”时才做格式校验
// =========================
                        if (!matchesExpectedType(cell, value, expectedType)) {
                            issues.add(new MergeIssue(
                                    file.getOriginalFilename(),
                                    sheet.getSheetName(),
                                    r + 1,
                                    template.headers().get(c),
                                    "格式与模板不一致"
                            ));
                        }

                    }
                    mergedRows.add(values);
                }
            } catch (Exception e) {
                issues.add(new MergeIssue(file.getOriginalFilename(), null, null, null,
                        "解析失败：" + e.getMessage()));
            }
        }

        mergedRowsRef.set(mergedRows);
        List<List<String>> preview = mergedRows.subList(0, Math.min(PREVIEW_LIMIT, mergedRows.size()));
        return new MergeResult(template.headers(), preview, mergedRows.size(), issues);
    }

    public byte[] exportMerged() {
        TemplateDefinition template = templateRef.get();
        List<List<String>> rows = mergedRowsRef.get();
        if (template == null || rows == null) {
            throw new IllegalStateException("没有可导出的汇总结果，请先完成合并。");
        }

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("汇总");
            Row header = sheet.createRow(0);
            for (int i = 0; i < template.headers().size(); i++) {
                header.createCell(i).setCellValue(template.headers().get(i));
            }

            for (int r = 0; r < rows.size(); r++) {
                Row row = sheet.createRow(r + 1);
                List<String> values = rows.get(r);
                for (int c = 0; c < values.size(); c++) {
                    row.createCell(c).setCellValue(values.get(c));
                }
            }

            for (int c = 0; c < template.headers().size(); c++) {
                sheet.autoSizeColumn(c);
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            return out.toByteArray();
        } catch (IOException e) {
            throw new IllegalStateException("导出失败：" + e.getMessage(), e);
        }
    }

    public class ExcelTemplateDefinition {
        private List<String> requiredColumns;  // 必填字段

        public List<String> getRequiredColumns() {
            return requiredColumns;
        }

        public void setRequiredColumns(List<String> requiredColumns) {
            this.requiredColumns = requiredColumns;
        }
    }




    // =========================
// 校验等级开关（默认 STRICT）
// =========================
    private enum ValidationLevel {
        STRICT,   // 严格：必填列为空 -> 报错
        LENIENT   // 宽松：空值不报错
    }

    // 👉 要的默认值：严格
    private final ValidationLevel validationLevel = ValidationLevel.STRICT;

//private final ValidationLevel validationLevel = ValidationLevel.LENIENT;

    private List<TemplateRule> loadTemplateRules() {
        ClassPathResource resource = new ClassPathResource("template-config.json");
        if (!resource.exists()) {
            return List.of();
        }
        ObjectMapper mapper = new ObjectMapper();
        try (InputStream input = resource.getInputStream()) {
            TemplateConfig config = mapper.readValue(input, TemplateConfig.class);
            if (config == null || config.templates() == null) {
                return List.of();
            }
            return config.templates().stream()
                    .filter(Objects::nonNull)
                    .toList();
        } catch (IOException e) {
            throw new IllegalStateException("模板配置读取失败：" + e.getMessage(), e);
        }
    }

    private Set<String> resolveRequiredHeaders(List<String> normalizedHeaders) {
        if (templateRules.isEmpty() || normalizedHeaders == null || normalizedHeaders.isEmpty()) {
            return Set.of();
        }
        Set<String> availableHeaders = new HashSet<>(normalizedHeaders);
        TemplateRule bestMatch = null;
        int bestScore = 0;
        for (TemplateRule rule : templateRules) {
            List<String> matchHeaders = normalizeHeaders(rule.matchHeaders());
            if (matchHeaders.isEmpty()) {
                continue;
            }
            if (availableHeaders.containsAll(matchHeaders)) {
                int score = matchHeaders.size();
                if (score > bestScore) {
                    bestMatch = rule;
                    bestScore = score;
                }
            }
        }
        if (bestMatch == null) {
            return Set.of();
        }
        return normalizeHeaders(bestMatch.requiredHeaders()).stream()
                .filter(availableHeaders::contains)
                .collect(LinkedHashSet::new, Set::add, Set::addAll);
    }

    private List<String> normalizeHeaders(List<String> headers) {
        if (headers == null || headers.isEmpty()) {
            return List.of();
        }
        List<String> normalized = new ArrayList<>();
        for (String header : headers) {
            String value = normalizeHeader(header);
            if (!value.isBlank()) {
                normalized.add(value);
            }
        }
        return normalized;
    }

    // ① 跳过被筛选隐藏的行（AutoFilter / 手动隐藏）
    private boolean isHiddenRow(Row row) {
        return row != null && row.getZeroHeight(); // 筛选隐藏/设置行高为0 时为 true
    }

    // ② 判断这一行是不是“真实数据行”
//    只填了序号不算；只要【除序号外】任意列有值，才算数据行
    private boolean isMeaningfulDataRow(Row row, DataFormatter fmt,
                                        List<String> coreHeaders,
                                        Set<String> requiredHeaders,
                                        Map<String, Integer> columnMap) {
        if (row == null) return false;
        if (row.getZeroHeight()) return false;

        if (requiredHeaders != null && !requiredHeaders.isEmpty()) {
            boolean hasMappedRequired = false;
            for (String required : requiredHeaders) {
                if (required == null || required.isBlank()) {
                    continue;
                }
                if (isIgnorableForRowDetection(required)) {
                    continue;
                }
                Integer col = columnMap.get(required);
                if (col == null) {
                    continue;
                }
                hasMappedRequired = true;
                Cell cell = row.getCell(col);
                String v = (cell == null) ? "" : fmt.formatCellValue(cell).trim();
                if (!v.isBlank()) {
                    return true;
                }
            }
            if (hasMappedRequired) {
                return false;
            }
        }

        boolean hasMappedCore = false;
        for (int i = 0; i < coreHeaders.size(); i++) {
            String norm = coreHeaders.get(i);
            if (norm == null) continue;
            if (isIgnorableForRowDetection(norm)) {
                continue;
            }
            Integer col = columnMap.get(norm);
            if (col == null) continue;
            hasMappedCore = true;

            Cell cell = row.getCell(col);
            String v = (cell == null) ? "" : fmt.formatCellValue(cell).trim();
            if (!v.isBlank()) {
                return true; // 只要有一个非序号字段有值，就认为是数据行
            }
        }
        if (hasMappedCore) {
            return false;
        }
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            String v = cell == null ? "" : fmt.formatCellValue(cell).trim();
            if (!v.isBlank()) {
                return true;
            }
        }
        return false;
    }

    private boolean isBusinessDataRow(Row row,
                                      DataFormatter fmt,
                                      Integer serialColumn,
                                      List<Integer> keyColumns,
                                      List<String> coreHeaders,
                                      Set<String> requiredHeaders,
                                      Map<String, Integer> columnMap) {
        if (row == null) {
            return false;
        }
        if (serialColumn != null) {
            Cell serialCell = row.getCell(serialColumn);
            if (!isValidSerialCell(serialCell, fmt)) {
                return false;
            }
            if (keyColumns != null && !keyColumns.isEmpty()) {
                for (Integer col : keyColumns) {
                    if (col == null) {
                        continue;
                    }
                    Cell cell = row.getCell(col);
                    String v = cell == null ? "" : fmt.formatCellValue(cell).trim();
                    if (!v.isBlank()) {
                        return true;
                    }
                }
                return false;
            }
        }
        return isMeaningfulDataRow(row, fmt, coreHeaders, requiredHeaders, columnMap);
    }

    private boolean isValidSerialCell(Cell cell, DataFormatter fmt) {
        if (cell == null) {
            return false;
        }
        String value = fmt.formatCellValue(cell).trim();
        if (value.isBlank()) {
            return false;
        }
        String normalized = value.replace(",", "");
        if (normalized.matches("\\d+")) {
            return true;
        }
        CellType cellType = cell.getCellType() == CellType.FORMULA
                ? cell.getCachedFormulaResultType()
                : cell.getCellType();
        if (cellType == CellType.NUMERIC && !DateUtil.isCellDateFormatted(cell)) {
            double numeric = cell.getNumericCellValue();
            return numeric == Math.floor(numeric);
        }
        return false;
    }

    private boolean shouldStopByInvalidRow(int invalidStreak) {
        return invalidStreak >= INVALID_ROW_LIMIT;
    }

    private Integer resolveSerialColumnIndex(TemplateDefinition template, Map<String, Integer> columnMap) {
        if (template != null) {
            for (String header : template.normalizedHeaders()) {
                if (isSerialHeader(header)) {
                    Integer col = columnMap.get(header);
                    if (col != null) {
                        return col;
                    }
                }
            }
        }
        for (Map.Entry<String, Integer> entry : columnMap.entrySet()) {
            if (isSerialHeader(entry.getKey())) {
                return entry.getValue();
            }
        }
        return null;
    }

    private List<Integer> resolveKeyColumns(TemplateDefinition template, Map<String, Integer> columnMap) {
        if (template == null || columnMap == null || columnMap.isEmpty()) {
            return List.of();
        }
        LinkedHashSet<Integer> indexes = new LinkedHashSet<>();
        for (String header : template.normalizedHeaders()) {
            if (header == null || header.isBlank()) {
                continue;
            }
            if (isSerialHeader(header)) {
                continue;
            }
            if (isKeyFieldHeader(header)) {
                Integer col = columnMap.get(header);
                if (col != null) {
                    indexes.add(col);
                }
            }
        }
        return new ArrayList<>(indexes);
    }

    private boolean isKeyFieldHeader(String normalizedHeader) {
        if (normalizedHeader == null || normalizedHeader.isBlank()) {
            return false;
        }
        for (String keyword : KEY_FIELD_KEYWORDS) {
            if (normalizedHeader.contains(keyword)) {
                return true;
            }
        }
        return false;
    }

    private List<String> resolveCoreHeaders(TemplateDefinition template) {
        if (template == null) {
            return List.of();
        }
        List<String> candidates = new ArrayList<>();
        Set<String> required = template.requiredNormalizedHeaders();
        if (required != null && !required.isEmpty()) {
            for (String header : required) {
                if (isCoreHeader(header)) {
                    candidates.add(header);
                }
            }
        }
        if (candidates.isEmpty()) {
            for (String header : template.normalizedHeaders()) {
                if (isCoreHeader(header)) {
                    candidates.add(header);
                }
            }
        }
        if (candidates.isEmpty()) {
            return template.normalizedHeaders();
        }
        return candidates;
    }

    private boolean isCoreHeader(String normalizedHeader) {
        if (normalizedHeader == null || normalizedHeader.isBlank()) {
            return false;
        }
        String header = normalizedHeader.trim();
        if (SERIAL_HEADER_PATTERN.matcher(header).matches()) {
            return false;
        }
        if (FIXED_VALUE_HEADER_PATTERN.matcher(header).matches()) {
            return false;
        }
        return !NON_CORE_HEADER_PATTERN.matcher(header).matches();
    }
    private boolean isSerialHeader(String normalizedHeader) {
        if (normalizedHeader == null || normalizedHeader.isBlank()) {
            return false;
        }
        return SERIAL_HEADER_PATTERN.matcher(normalizedHeader.trim()).matches();
    }

    private boolean isIgnorableForRowDetection(String normalizedHeader) {
        if (normalizedHeader == null || normalizedHeader.isBlank()) {
            return true;
        }
        String header = normalizedHeader.trim();
        return isSerialHeader(header) || FIXED_VALUE_HEADER_PATTERN.matcher(header).matches();
    }

    // =========================
    // ✅ 表头定位：改进版
    // =========================

    private int findHeaderRowByDensity(Sheet sheet) {
        int first = sheet.getFirstRowNum();
        int last = Math.min(sheet.getLastRowNum(), first + HEADER_SCAN_LIMIT);

        int bestRow = -1;
        int bestTextCount = 0;
        int bestNonEmptyCount = 0;
        int instructionRowFallback = -1;

        DataFormatter fmt = new DataFormatter();

        for (int r = first; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }

            int nonEmptyCount = 0;
            int textCount = 0;
            int firstNonEmptyCol = -1;
            String mainText = null;

            for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                String value = cell == null ? "" : fmt.formatCellValue(cell).trim();
                if (value.isBlank()) {
                    continue;
                }
                if (firstNonEmptyCol < 0) {
                    firstNonEmptyCol = c;
                    mainText = value;
                }
                nonEmptyCount++;
                if (isHeaderTextCell(cell, value)) {
                    textCount++;
                }
            }

            if (textCount == 0 && nonEmptyCount == 0) {
                continue;
            }

            // 合并单元格标题说明行
            if (isInstructionRow(sheet, r, firstNonEmptyCol, nonEmptyCount)) {
                if (instructionRowFallback < 0) {
                    instructionRowFallback = r;
                }
                continue;
            }

            // 非合并单元格说明行：只有一个有效格 + 命中关键词
            if (nonEmptyCount == 1 && looksLikeInstructionText(mainText)) {
                // 优先尝试下一行
                int next = r + 1;
                if (next <= sheet.getLastRowNum()) {
                    Row nextRow = sheet.getRow(next);
                    if (isLikelyHeaderRow(nextRow, fmt)) {
                        return next;
                    }
                }
                if (instructionRowFallback < 0) {
                    instructionRowFallback = r;
                }
                continue;
            }

            if (textCount > bestTextCount || (textCount == bestTextCount && nonEmptyCount > bestNonEmptyCount)) {
                bestTextCount = textCount;
                bestNonEmptyCount = nonEmptyCount;
                bestRow = r;
            }
        }

        if (bestRow >= 0) {
            if (bestTextCount == 0) {
                return bestNonEmptyCount == 0 ? -1 : bestRow;
            }
            return bestRow;
        }
        return instructionRowFallback;
    }

    private int findFirstNonEmptyRow(Sheet sheet) {
        int first = sheet.getFirstRowNum();
        int last = Math.min(sheet.getLastRowNum(), first + HEADER_SCAN_LIMIT);
        DataFormatter fmt = new DataFormatter();

        for (int r = first; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                String value = cell == null ? "" : fmt.formatCellValue(cell).trim();
                if (!value.isBlank()) {
                    return r;
                }
            }
        }
        return -1;
    }

    private int findHeaderRowByMatch(Sheet sheet, List<String> templateHeaders) {
        int first = sheet.getFirstRowNum();
        int last = Math.min(sheet.getLastRowNum(), first + HEADER_SCAN_LIMIT);

        DataFormatter fmt = new DataFormatter();

        for (int r = first; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }

            // ✅ 跳过说明行
            if (isInstructionLikeRow(row, fmt)) {
                continue;
            }

            List<String> rowHeaders = new ArrayList<>();
            for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                String value = cell == null ? "" : normalizeHeader(fmt.formatCellValue(cell));
                if (!value.isBlank()) {
                    rowHeaders.add(value);
                }
            }
            if (isExactHeaderMatch(rowHeaders, templateHeaders)) {
                return r;
            }
        }
        return -1;
    }


    private boolean isInstructionLikeRow(Row row, DataFormatter fmt) {
        if (row == null) return false;

        int nonEmpty = 0;
        String main = null;

        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            String v = cell == null ? "" : fmt.formatCellValue(cell).trim();
            if (v.isBlank()) continue;
            nonEmpty++;
            if (main == null) main = v;
        }
        return nonEmpty == 1 && looksLikeInstructionText(main);
    }

    private boolean isLikelyHeaderRow(Row row, DataFormatter fmt) {
        if (row == null) return false;

        int nonEmpty = 0;
        int text = 0;

        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            String v = cell == null ? "" : fmt.formatCellValue(cell).trim();
            if (v.isBlank()) continue;

            if (looksLikeInstructionText(v)) return false;

            nonEmpty++;
            if (isHeaderTextCell(cell, v)) {
                text++;
            }
        }

        if (nonEmpty < 2) return false;
        return text >= Math.max(2, (int) Math.ceil(nonEmpty * 0.6));
    }

    private boolean looksLikeInstructionText(String v) {
        if (v == null) return false;
        String s = v.trim();
        if (s.isBlank()) return false;
        return INSTRUCTION_KEYWORDS.matcher(s).matches();
    }

    // =========================
    // 原有逻辑保持不动
    // =========================

    private ColumnMapping buildColumnMapping(Row headerRow) {
        Map<String, Integer> map = new LinkedHashMap<>();
        Set<String> duplicates = new LinkedHashSet<>();
        if (headerRow == null) {
            return new ColumnMapping(map, duplicates);
        }
        DataFormatter fmt = new DataFormatter();
        for (int c = headerRow.getFirstCellNum(); c < headerRow.getLastCellNum(); c++) {
            Cell cell = headerRow.getCell(c);
            String name = cell == null ? "" : fmt.formatCellValue(cell).trim();
            if (name.isBlank()) {
                continue;
            }
            String normalized = normalizeHeader(name);
            if (normalized.isBlank()) {
                continue;
            }
            if (map.containsKey(normalized)) {
                duplicates.add(normalized);
                continue;
            }
            map.put(normalized, c);
        }
        return new ColumnMapping(map, duplicates);
    }

    private List<ColumnType> detectColumnTypes(Sheet sheet, int dataStartRow, List<Integer> columnIndexes) {
        List<ColumnType> types = new ArrayList<>();
        for (Integer col : columnIndexes) {
            types.add(detectColumnTypeForColumn(sheet, dataStartRow, col));
        }
        return types;
    }

    private ColumnType detectColumnTypeForColumn(Sheet sheet, int dataStartRow, int columnIndex) {
        int last = Math.min(sheet.getLastRowNum(), dataStartRow + TYPE_SAMPLE_LIMIT);
        for (int r = dataStartRow; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            Cell cell = row.getCell(columnIndex);
            if (cell == null) {
                continue;
            }
            CellType cellType = cell.getCellType();
            if (cellType == CellType.FORMULA) {
                cellType = cell.getCachedFormulaResultType();
            }
            if (cellType == CellType.NUMERIC) {
                return DateUtil.isCellDateFormatted(cell) ? ColumnType.DATE : ColumnType.NUMBER;
            }
            if (cellType == CellType.STRING || cellType == CellType.BOOLEAN) {
                return ColumnType.TEXT;
            }
        }
        return ColumnType.TEXT;
    }

    private boolean matchesExpectedType(Cell cell, String value, ColumnType expectedType) {
        if (expectedType == ColumnType.TEXT) {
            return true;
        }
        if (expectedType == ColumnType.NUMBER) {
            if (cell != null) {
                CellType cellType = cell.getCellType() == CellType.FORMULA
                        ? cell.getCachedFormulaResultType()
                        : cell.getCellType();
                if (cellType == CellType.NUMERIC) {
                    return true;
                }
            }
            return isNumeric(value);
        }
        if (expectedType == ColumnType.DATE) {
            if (cell != null) {
                CellType cellType = cell.getCellType() == CellType.FORMULA
                        ? cell.getCachedFormulaResultType()
                        : cell.getCellType();
                if (cellType == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                    return true;
                }
            }
            return isDateString(value);
        }
        return true;
    }

    private boolean isRowBlank(Row row, DataFormatter fmt) {
        if (row == null) {
            return true;
        }
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            String v = cell == null ? "" : fmt.formatCellValue(cell);
            if (v != null && !v.trim().isBlank()) {
                return false;
            }
        }
        return true;
    }

    private boolean isNumeric(String value) {
        try {
            Double.parseDouble(value.replace(",", ""));
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private boolean isDateString(String value) {
        String normalized = value.trim().replace('.', '-').replace('/', '-');
        List<DateTimeFormatter> formats = List.of(
                DateTimeFormatter.ofPattern("yyyyMMdd"),
                DateTimeFormatter.ofPattern("yyyy-M-d"),
                DateTimeFormatter.ofPattern("yyyy-MM-dd"),
                DateTimeFormatter.ofPattern("yyyy年M月d日")
        );
        for (DateTimeFormatter fmt : formats) {
            try {
                LocalDate.parse(normalized, fmt);
                return true;
            } catch (Exception ignored) {
            }
        }
        return false;
    }

    private String normalizeHeader(String raw) {
        if (raw == null) {
            return "";
        }
        String s = raw.trim();
        if (s.isBlank()) {
            return "";
        }
        s = s.replace("\n", "").replace("\r", "").trim();
        s = s.replaceAll("（.*?）", "");
        s = s.replaceAll("\\(.*?\\)", "");
        s = s.replaceAll("\\*", "");
        s = s.replaceAll("\\s+", "");
        return s.trim();
    }
    private String resolveHeaderName(TemplateDefinition template, String normalizedHeader) {
        if (template == null || normalizedHeader == null) {
            return normalizedHeader;
        }
        List<String> normalizedHeaders = template.normalizedHeaders();
        for (int i = 0; i < normalizedHeaders.size(); i++) {
            if (normalizedHeader.equals(normalizedHeaders.get(i))) {
                return template.headers().get(i);
            }
        }
        return normalizedHeader;
    }

    private record ColumnMapping(Map<String, Integer> columnMap, Set<String> duplicateHeaders) {
    }

    private boolean isHeaderTextCell(Cell cell, String value) {
        if (cell != null) {
            CellType cellType = cell.getCellType() == CellType.FORMULA
                    ? cell.getCachedFormulaResultType()
                    : cell.getCellType();
            if (cellType == CellType.STRING) {
                return true;
            }
        }
        if (value == null) {
            return false;
        }
        String trimmed = value.trim();
        if (trimmed.isBlank()) {
            return false;
        }
        return HEADER_TEXT_PATTERN.matcher(trimmed).matches();
    }

    private boolean isInstructionRow(Sheet sheet, int rowIndex, int firstNonEmptyCol, int nonEmptyCount) {
        if (nonEmptyCount != 1 || firstNonEmptyCol < 0) {
            return false;
        }
        int mergedCount = sheet.getNumMergedRegions();
        if (mergedCount == 0) {
            return false;
        }
        for (int i = 0; i < mergedCount; i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.getFirstRow() <= rowIndex && region.getLastRow() >= rowIndex
                    && region.getFirstColumn() <= firstNonEmptyCol && region.getLastColumn() >= firstNonEmptyCol) {
                return region.getLastColumn() > region.getFirstColumn();
            }
        }
        return false;
    }

    private boolean isExactHeaderMatch(List<String> rowHeaders, List<String> templateHeaders) {
        if (rowHeaders == null || templateHeaders == null) {
            return false;
        }
        List<String> filteredRow = rowHeaders.stream()
                .filter(v -> v != null && !v.isBlank())
                .toList();
        if (filteredRow.isEmpty() || templateHeaders.isEmpty()) {
            return false;
        }
        if (filteredRow.size() != templateHeaders.size()) {
            return false;
        }
        Set<String> rowSet = new LinkedHashSet<>(filteredRow);
        if (rowSet.size() != filteredRow.size()) {
            return false;
        }
        return filteredRow.equals(templateHeaders);
    }
    private Sheet pickBestDataSheet(Workbook workbook) {
        DataFormatter fmt = new DataFormatter();

        Sheet best = null;
        int bestScore = -1;

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet == null) continue;

            int score = 0;
            int maxRow = Math.min(sheet.getLastRowNum(), 80); // 只看前80行即可
            for (int r = sheet.getFirstRowNum(); r <= maxRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                // 只看前50列防止超宽表浪费
                short firstCell = row.getFirstCellNum();
                short lastCell = row.getLastCellNum();
                if (firstCell < 0 || lastCell < 0) continue;

                int endCol = Math.min(lastCell, (short) (firstCell + 50));
                for (int c = firstCell; c < endCol; c++) {
                    Cell cell = row.getCell(c);
                    String v = cell == null ? "" : fmt.formatCellValue(cell).trim();
                    if (!v.isBlank()) score++;
                }
            }

            // 至少要有一点内容才算数据sheet
            if (score > bestScore) {
                bestScore = score;
                best = sheet;
            }
        }

        // 兜底：全都空就返回第一个
        return best != null ? best : workbook.getSheetAt(0);
    }


}
