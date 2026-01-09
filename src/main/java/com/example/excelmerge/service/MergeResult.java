package com.example.excelmerge.service;

import java.util.List;

public record MergeResult(
        List<String> headers,
        List<List<String>> previewRows,
        int totalRows,
        List<MergeIssue> issues
) {
}