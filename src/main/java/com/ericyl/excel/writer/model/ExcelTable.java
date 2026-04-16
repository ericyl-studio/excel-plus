package com.ericyl.excel.writer.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;

import java.util.List;

@Data
@Accessors(chain = true)
@SuperBuilder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelTable {
    private List<List<ExcelColumn>> headers;
    private List<List<ExcelColumn>> columns;
    private List<List<ExcelColumn>> footers;
}
