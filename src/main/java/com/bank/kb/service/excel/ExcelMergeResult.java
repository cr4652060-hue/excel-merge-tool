package com.bank.kb.service.excel;

import java.util.List;

public record ExcelMergeResult(
        List<String> headers,
        List<List<String>> previewRows,
        int totalRows,
        List<ExcelMergeIssue> issues
) {
}