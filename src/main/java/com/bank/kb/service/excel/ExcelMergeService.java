package com.bank.kb.service.excel;

import com.example.excelmerge.service.MergeIssue;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.annotation.Value;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;
import java.util.regex.Pattern;

@Service
public class ExcelMergeService {
    private static final int INVALID_ROW_LIMIT = 30;
    private static final int HEADER_SCAN_LIMIT = 30;
    private static final int TYPE_SAMPLE_LIMIT = 50;
    private static final int PREVIEW_LIMIT = 500;

    private static final Pattern HEADER_TEXT_PATTERN = Pattern.compile(".*[A-Za-z\\u4e00-\\u9fff].*");
    private static final Pattern SERIAL_HEADER_PATTERN = Pattern.compile("^(序号|序|编号|行号|序列|no|No|NO)$");
    private static final Pattern FIXED_VALUE_HEADER_PATTERN = Pattern.compile(".*(账户类型|账户类别).*");
    // ✅ 新增：说明行/标题行关键词（内网填报模板常见话术）
    private static final Pattern INSTRUCTION_KEYWORDS = Pattern.compile(
            ".*(填写|说明|注意|示例|要求|口径|备注|提示|温馨提示|如实|以下|请按|请填写|填报|填表|规则|校验|检查).*"
    );
    private static final String ANCHOR_KEYWORDS_PROPERTY = "excel.merge.keywords.anchors";
    private static final String KEY_FIELD_KEYWORDS_PROPERTY = "excel.merge.keywords.keys";
    private static final String EXCLUDED_KEYWORDS_PROPERTY = "excel.merge.keywords.excludes";
    private static final String TOTAL_KEYWORDS_PROPERTY = "excel.merge.keywords.totals";
    private static final String KEY_FIELD_MIN_HITS_PROPERTY = "excel.merge.keywords.minHits";

    private static final List<String> DEFAULT_ANCHOR_KEYWORDS = List.of(
            "账号",
            "卡号",
            "证件号",
            "设备序列号",
            "资产编号",
            "设备编号"
    );
    private static final List<String> DEFAULT_KEY_FIELD_KEYWORDS = List.of(
            "姓名",
            "单位",
            "网点",
            "部门",
            "金额",
            "数量",
            "用途",
            "存放地点",
            "管理员",
            "项目",
            "指标",
            "设备类型",
            "规格型号",
            "设备名称",
            "资产名称"
    );
    private static final List<String> DEFAULT_EXCLUDED_KEYWORDS = List.of(
            "序号",
            "序次",
            "行号",
            "备注",
            "说明",
            "填报人",
            "填表人",
            "填报日期",
            "填表日期"
    );
    private static final List<String> DEFAULT_TOTAL_KEYWORDS = List.of("小计", "合计", "总计");
    private static final int DEFAULT_KEY_FIELD_MIN_HITS = 2;
    private final AtomicReference<ExcelTemplateDefinition> templateRef = new AtomicReference<>();
    private final AtomicReference<List<List<String>>> mergedRowsRef = new AtomicReference<>();


    public ExcelTemplateInfo analyzeTemplate(MultipartFile file) {
        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = pickBestDataSheet(workbook);

            int headerRow = findHeaderRowByDensity(sheet);
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
            ExcelTemplateDefinition definition = new ExcelTemplateDefinition(
                    headers,
                    normalized,
                    types,
                    headerRow,
                    headerRow + 1
            );
            templateRef.set(definition);
            mergedRowsRef.set(null);

            return new ExcelTemplateInfo(headers, headerRow + 1, headerRow + 2, types);
        } catch (IOException e) {
            throw new IllegalStateException("模板解析失败：" + e.getMessage(), e);
        }
    }

    public ExcelMergeResult mergeFiles(List<MultipartFile> files) {
        ExcelTemplateDefinition template = templateRef.get();
        if (template == null) {
            throw new IllegalStateException("请先上传模板文件，再进行合并。");
        }
        if (files == null || files.isEmpty()) {
            throw new IllegalArgumentException("请至少上传一份支行 Excel。");
        }

        List<List<String>> mergedRows = new ArrayList<>();
        List<ExcelMergeIssue> issues = new ArrayList<>();

        for (MultipartFile file : files) {
            if (file.isEmpty()) {
                issues.add(new ExcelMergeIssue(file.getOriginalFilename(), null, null, null, "文件为空，已跳过。"));
                continue;
            }
            try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
                Sheet sheet = pickBestDataSheet(workbook);

                // ✅ 改进：匹配模板表头时，也跳过说明行/标题行
                int headerRowIndex = findHeaderRowByMatch(sheet, template.normalizedHeaders());
                if (headerRowIndex < 0) {
                    issues.add(new ExcelMergeIssue(file.getOriginalFilename(), sheet.getSheetName(), null, null,
                            "未找到匹配模板的表头，已跳过。"));
                    continue;
                }

                Map<String, Integer> columnMap = buildColumnMap(sheet.getRow(headerRowIndex));
                Set<String> missingColumns = new HashSet<>();
                for (int i = 0; i < template.normalizedHeaders().size(); i++) {
                    String norm = template.normalizedHeaders().get(i);
                    if (!columnMap.containsKey(norm)) {
                        missingColumns.add(norm);
                        issues.add(new ExcelMergeIssue(file.getOriginalFilename(), sheet.getSheetName(), null,
                                template.headers().get(i), "缺少列：" + template.headers().get(i)));
                    }
                }

                int lastRow = sheet.getLastRowNum();
                DataFormatter fmt = new DataFormatter();
                List<Integer> editableColumns = resolveEditableColumnIndexes(template, columnMap);
                KeyColumnInfo keyColumnInfo = resolveKeyColumnInfo(template, columnMap);
                int emptyEditableStreak = 0;
                for (int r = headerRowIndex + 1; r <= lastRow; r++) {
                    Row row = sheet.getRow(r);
                    // ✅ 1) 跳过空行
                    if (row == null) {
                        if (shouldStopByInvalidRow(++emptyEditableStreak)) {
                            break;
                        }
                        continue;
                    }

                    // ✅ 2) 跳过筛选隐藏行（只合并“可见行”）
                    if (isHiddenRow(row)) {
                        if (shouldStopByInvalidRow(++emptyEditableStreak)) {
                            break;
                        }
                        continue;
                    }

                    // ✅ 3) 关键字段命中判定数据行（默认规则 + 可配置关键词）
                    if (isTotalRow(row, fmt)) {
                        break;
                    }
                    if (!isDataRowByKeyColumns(row, fmt, keyColumnInfo, editableColumns)) {
                        if (shouldStopByInvalidRow(++emptyEditableStreak)) {
                            break;
                        }
                        continue;
                    }
                    emptyEditableStreak = 0;
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
                                    && isRequiredHeader(template.headers().get(c))) {

                                issues.add(new ExcelMergeIssue(
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
                            issues.add(new ExcelMergeIssue(
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
                issues.add(new ExcelMergeIssue(file.getOriginalFilename(), null, null, null,
                        "解析失败：" + e.getMessage()));
            }
        }

        mergedRowsRef.set(mergedRows);
        List<List<String>> preview = mergedRows.subList(0, Math.min(PREVIEW_LIMIT, mergedRows.size()));
        return new ExcelMergeResult(template.headers(), preview, mergedRows.size(), issues);
    }

    public byte[] exportMerged() {
        ExcelTemplateDefinition template = templateRef.get();
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





    // =========================
// 校验等级开关（默认 STRICT）
// =========================
    private enum ValidationLevel {
        STRICT,   // 严格：必填列为空 -> 报错
        LENIENT   // 宽松：空值不报错
    }

    // 👉 要的默认值：严格
    private final ValidationLevel validationLevel = ValidationLevel.STRICT;

    //想“先能合并就行”：
//private final ValidationLevel validationLevel = ValidationLevel.LENIENT;

    private boolean isRequiredHeader(String header) {
        if (header == null) return false;
        String h = header.replaceAll("\\s+", "");

        // 金额、备注：允许为空
        if (h.contains("金额") || h.contains("备注")) return false;

        // 必填项（按你们网点表结构）
        return h.contains("单位") || h.contains("网点")
                || h.contains("账号") || h.contains("卡号")
                || h.contains("姓名")
                || h.contains("账户类型") || h.contains("账户类别");
    }
    // ① 跳过被筛选隐藏的行（AutoFilter / 手动隐藏）
    private boolean isHiddenRow(Row row) {
        return row != null && row.getZeroHeight(); // 筛选隐藏/设置行高为0 时为 true
    }

    private boolean hasEditableValue(Row row, DataFormatter fmt, List<Integer> editableColumns) {
        if (row == null || editableColumns == null || editableColumns.isEmpty()) {
            return false;
        }
        for (Integer col : editableColumns) {
            if (col == null) {
                continue;
            }

            Cell cell = row.getCell(col);
            String value = cell == null ? "" : fmt.formatCellValue(cell).trim();
            if (!value.isBlank()) {
                return true;
            }
        }
        return false;
    }
    private boolean shouldStopByInvalidRow(int invalidStreak) {
        return invalidStreak >= INVALID_ROW_LIMIT;
    }

    private List<Integer> resolveEditableColumnIndexes(ExcelTemplateDefinition template,
                                                       Map<String, Integer> columnMap) {
        if (template == null || columnMap == null || columnMap.isEmpty()) {
            return List.of();
        }
        LinkedHashSet<Integer> indexes = new LinkedHashSet<>();
        for (String header : template.normalizedHeaders()) {
            if (header == null || header.isBlank()) {
                continue;
            }
            if (isIgnorableForRowDetection(header)) {
                continue;
            }
            Integer col = columnMap.get(header);
            if (col != null) {
                indexes.add(col);
            }
        }
        if (indexes.isEmpty()) {
            for (String header : template.normalizedHeaders()) {
                Integer col = columnMap.get(header);
                if (col != null) {
                    indexes.add(col);
                }
            }
        }
        return new ArrayList<>(indexes);
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
    private boolean isExcludedHeaderForRowDetection(String normalizedHeader, List<String> excludedKeywords) {
        if (normalizedHeader == null || normalizedHeader.isBlank()) {
            return true;
        }
        String header = normalizedHeader.trim();
        if (isSerialHeader(header)) {
            return true;
        }
        if (FIXED_VALUE_HEADER_PATTERN.matcher(header).matches()) {
            return true;
        }
        return containsKeyword(header, excludedKeywords);
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

            // ✅ 1) 合并单元格标题说明行（你原来的逻辑保留）
            if (isInstructionRow(sheet, r, firstNonEmptyCol, nonEmptyCount)) {
                if (instructionRowFallback < 0) {
                    instructionRowFallback = r;
                }
                continue;
            }

            // ✅ 2) 非合并单元格的说明行：只有一个有效格 + 命中“填写说明/注意/口径”等关键词
            if (nonEmptyCount == 1 && looksLikeInstructionText(mainText)) {
                // 如果下一行更像表头：直接选下一行
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

            // ✅ 3) 普通评分选最优
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

    private int findHeaderRowByMatch(Sheet sheet, List<String> templateHeaders) {
        int first = sheet.getFirstRowNum();
        int last = Math.min(sheet.getLastRowNum(), first + HEADER_SCAN_LIMIT);

        int bestRow = -1;
        int bestCount = 0;
        DataFormatter fmt = new DataFormatter();

        for (int r = first; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }

            // ✅ 跳过“说明行”（防止说明里含字段示例导致误命中）
            if (isInstructionLikeRow(row, fmt)) {
                continue;
            }

            int count = 0;
            for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                String value = cell == null ? "" : normalizeHeader(fmt.formatCellValue(cell));
                if (!value.isBlank() && templateHeaders.contains(value)) {
                    count++;
                }
            }
            if (count > bestCount) {
                bestCount = count;
                bestRow = r;
            }
        }
        return bestCount == 0 ? -1 : bestRow;
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

            // 表头一般不出现“填写说明/注意事项”
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
    // 下面保持原有代码不动
    // =========================

    private Map<String, Integer> buildColumnMap(Row headerRow) {
        Map<String, Integer> map = new LinkedHashMap<>();
        if (headerRow == null) {
            return map;
        }
        DataFormatter fmt = new DataFormatter();
        for (int c = headerRow.getFirstCellNum(); c < headerRow.getLastCellNum(); c++) {
            Cell cell = headerRow.getCell(c);
            String name = cell == null ? "" : fmt.formatCellValue(cell).trim();
            if (name.isBlank()) {
                continue;
            }
            map.put(normalizeHeader(name), c);
        }
        return map;
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
                DateTimeFormatter.ofPattern("yyyy-M-d"),
                DateTimeFormatter.ofPattern("yyyy-MM-dd")
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

    private KeyColumnInfo resolveKeyColumnInfo(ExcelTemplateDefinition template, Map<String, Integer> columnMap) {
        if (template == null || columnMap == null || columnMap.isEmpty()) {
            return new KeyColumnInfo(List.of(), List.of(), DEFAULT_KEY_FIELD_MIN_HITS);
        }
        List<String> anchorKeywords = loadKeywords(ANCHOR_KEYWORDS_PROPERTY, DEFAULT_ANCHOR_KEYWORDS);
        List<String> keyKeywords = loadKeywords(KEY_FIELD_KEYWORDS_PROPERTY, DEFAULT_KEY_FIELD_KEYWORDS);
        List<String> excludedKeywords = loadKeywords(EXCLUDED_KEYWORDS_PROPERTY, DEFAULT_EXCLUDED_KEYWORDS);
        int minHits = loadMinKeyHits();

        LinkedHashSet<Integer> anchorIndexes = new LinkedHashSet<>();
        LinkedHashSet<Integer> keyIndexes = new LinkedHashSet<>();
        for (String header : template.normalizedHeaders()) {
            if (header == null || header.isBlank()) {
                continue;
            }
            if (isExcludedHeaderForRowDetection(header, excludedKeywords)) {
                continue;
            }
            Integer col = columnMap.get(header);
            if (col == null) {
                continue;
            }
            if (containsKeyword(header, anchorKeywords)) {
                anchorIndexes.add(col);
                continue;
            }
            if (containsKeyword(header, keyKeywords)) {
                keyIndexes.add(col);
            }
        }
        return new KeyColumnInfo(new ArrayList<>(anchorIndexes), new ArrayList<>(keyIndexes), minHits);
    }

    private boolean isDataRowByKeyColumns(Row row,
                                          DataFormatter fmt,
                                          KeyColumnInfo keyColumnInfo,
                                          List<Integer> editableColumns) {
        if (row == null) {
            return false;
        }
        if (keyColumnInfo == null) {
            return hasEditableValue(row, fmt, editableColumns);
        }
        List<Integer> anchorColumns = keyColumnInfo.anchorColumns();
        if (anchorColumns != null && !anchorColumns.isEmpty()) {
            for (Integer col : anchorColumns) {
                if (col == null) {
                    continue;
                }
                Cell cell = row.getCell(col);
                String value = cell == null ? "" : fmt.formatCellValue(cell).trim();
                if (!value.isBlank()) {
                    return true;
                }
            }
        }
        List<Integer> keyColumns = keyColumnInfo.keyColumns();
        if (keyColumns != null && !keyColumns.isEmpty()) {
            int hits = 0;
            for (Integer col : keyColumns) {
                if (col == null) {
                    continue;
                }
                Cell cell = row.getCell(col);
                String value = cell == null ? "" : fmt.formatCellValue(cell).trim();
                if (!value.isBlank()) {
                    hits++;
                }
            }
            return hits >= keyColumnInfo.minHits();
        }
        return hasEditableValue(row, fmt, editableColumns);
    }

    private boolean isTotalRow(Row row, DataFormatter fmt) {
        if (row == null) {
            return false;
        }
        List<String> totals = loadKeywords(TOTAL_KEYWORDS_PROPERTY, DEFAULT_TOTAL_KEYWORDS);
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            String value = cell == null ? "" : fmt.formatCellValue(cell).trim();
            if (value.isBlank()) {
                continue;
            }
            if (containsKeyword(value, totals)) {
                return true;
            }
        }
        return false;
    }

    private List<String> loadKeywords(String propertyName, List<String> defaults) {
        String raw = System.getProperty(propertyName);
        if (raw == null || raw.isBlank()) {
            return defaults;
        }
        String[] parts = raw.split("[,，;；]");
        List<String> values = new ArrayList<>();
        for (String part : parts) {
            String trimmed = part == null ? "" : part.trim();
            if (!trimmed.isBlank()) {
                values.add(trimmed);
            }
        }
        return values.isEmpty() ? defaults : values;
    }

    private int loadMinKeyHits() {
        String raw = System.getProperty(KEY_FIELD_MIN_HITS_PROPERTY);
        if (raw == null || raw.isBlank()) {
            return DEFAULT_KEY_FIELD_MIN_HITS;
        }
        try {
            int value = Integer.parseInt(raw.trim());
            return Math.max(1, value);
        } catch (NumberFormatException e) {
            return DEFAULT_KEY_FIELD_MIN_HITS;
        }
    }

    private boolean containsKeyword(String header, List<String> keywords) {
        if (header == null || header.isBlank() || keywords == null || keywords.isEmpty()) {
            return false;
        }
        for (String keyword : keywords) {
            if (keyword == null || keyword.isBlank()) {
                continue;
            }
            if (header.contains(keyword)) {
                return true;
            }
        }
        return false;
    }

    private record KeyColumnInfo(List<Integer> anchorColumns,
                                 List<Integer> keyColumns,
                                 int minHits) {
    }




}
