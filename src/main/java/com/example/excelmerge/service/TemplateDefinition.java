package com.example.excelmerge.service;

import java.util.List;

public record TemplateDefinition(
        List<String> headers,
        List<String> normalizedHeaders,
        List<ColumnType> columnTypes,
        int headerRowIndex,
        int dataStartRow
) {
}