package com.ericyl.excel.writer.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.util.CellRangeAddress;

@Data
@AllArgsConstructor
public class ExcelRegion {

    private CellRangeAddress region;

    private ExcelColumn excelColumn;

}
