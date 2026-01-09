package com.example.excelmerge.service;

import java.util.List;

public record TemplateInfo(
        List<String> headers,
        int headerRowIndex,
        int dataStartRow,
        List<ColumnType> columnTypes
) {
}