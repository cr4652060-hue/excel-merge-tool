package com.example.excelmerge.service;

import java.util.List;
import java.util.Set;

public record TemplateDefinition(
        List<String> headers,
        List<String> normalizedHeaders,
        List<ColumnType> columnTypes,
        Set<String> requiredNormalizedHeaders,
        Set<String> editableNormalizedHeaders,
        int headerRowIndex,
        int dataStartRow
) {
}