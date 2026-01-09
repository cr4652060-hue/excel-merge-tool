package com.example.excelmerge.service;

public record MergeIssue(
        String fileName,
        String sheetName,
        Integer rowNo,
        String columnName,
        String message
) {
}