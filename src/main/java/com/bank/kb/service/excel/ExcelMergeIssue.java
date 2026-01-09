package com.bank.kb.service.excel;

public record ExcelMergeIssue(
        String fileName,
        String sheetName,
        Integer rowNo,
        String columnName,
        String message
) {
}