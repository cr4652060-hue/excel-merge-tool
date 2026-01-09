package com.bank.kb.service.excel;

import java.util.List;

public record ExcelTemplateDefinition(
        List<String> headers,
        List<String> normalizedHeaders,
        List<ColumnType> columnTypes,
        int headerRowIndex,
        int dataStartRow
) {
}