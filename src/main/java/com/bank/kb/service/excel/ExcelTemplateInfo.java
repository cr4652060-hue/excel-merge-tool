package com.bank.kb.service.excel;

import java.util.List;

public record ExcelTemplateInfo(
        List<String> headers,
        int headerRowIndex,
        int dataStartRow,
        List<ColumnType> columnTypes
) {
}