package com.bank.kb.service.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;
import java.util.regex.Pattern;
@Service
public class ExcelMergeService {
    private static final int HEADER_SCAN_LIMIT = 30;
    private static final int TYPE_SAMPLE_LIMIT = 50;
    private static final int PREVIEW_LIMIT = 50;
    private static final Pattern HEADER_TEXT_PATTERN = Pattern.compile(".*[A-Za-z\\u4e00-\\u9fff].*");

    private final AtomicReference<ExcelTemplateDefinition> templateRef = new AtomicReference<>();
    private final AtomicReference<List<List<String>>> mergedRowsRef = new AtomicReference<>();

    public ExcelTemplateInfo analyzeTemplate(MultipartFile file) {
        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);
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
                Sheet sheet = workbook.getSheetAt(0);
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
                for (int r = headerRowIndex + 1; r <= lastRow; r++) {
                    Row row = sheet.getRow(r);
                    if (isRowBlank(row, fmt)) {
                        continue;
                    }
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
                        if (value.isBlank()) {
                            issues.add(new ExcelMergeIssue(file.getOriginalFilename(), sheet.getSheetName(),
                                    r + 1, template.headers().get(c), "单元格为空"));
                            continue;
                        }

                        if (!matchesExpectedType(cell, value, expectedType)) {
                            issues.add(new ExcelMergeIssue(file.getOriginalFilename(), sheet.getSheetName(),
                                    r + 1, template.headers().get(c), "格式与模板不一致"));
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

    private int findHeaderRowByDensity(Sheet sheet) {
        int first = sheet.getFirstRowNum();
        int last = Math.min(sheet.getLastRowNum(), first + HEADER_SCAN_LIMIT);
        int bestRow = -1;
        int bestTextCount = 0;
        int bestNonEmptyCount = 0;
        DataFormatter fmt = new DataFormatter();
        for (int r = first; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            int nonEmptyCount = 0;
            int textCount = 0;
            for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                String value = cell == null ? "" : fmt.formatCellValue(cell).trim();
                if (value.isBlank()) {
                    continue;
                }
                nonEmptyCount++;
                if (isHeaderTextCell(cell, value)) {
                    textCount++;
                }
            }
            if (textCount == 0 && nonEmptyCount == 0) {
                continue;
            }
            if (textCount > bestTextCount || (textCount == bestTextCount && nonEmptyCount > bestNonEmptyCount)) {
                bestTextCount = textCount;
                bestNonEmptyCount = nonEmptyCount;
                bestRow = r;
            }
        }
        if (bestRow < 0) {
            return -1;
        }
        if (bestTextCount == 0) {
            return bestNonEmptyCount == 0 ? -1 : bestRow;
        }
        return bestRow;
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
}