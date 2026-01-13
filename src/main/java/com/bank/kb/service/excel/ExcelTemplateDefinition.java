package com.bank.kb.service.excel;

import java.util.List;
public record ExcelTemplateDefinition(
        List<String> headers,          // 表头
        List<String> normalizedHeaders,  // 归一化表头
        List<ColumnType> columnTypes,   // 列类型
        int headerRowIndex,            // 表头所在行
        int dataStartRow               // 数据开始行
) {

}
