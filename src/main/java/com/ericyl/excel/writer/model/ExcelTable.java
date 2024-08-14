package com.ericyl.excel.writer.model;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.List;

@Data
@AllArgsConstructor
public class ExcelTable {
    private List<List<ExcelColumn>> headers;
    private List<List<ExcelColumn>> columns;
    private List<List<ExcelColumn>> footers;
}
