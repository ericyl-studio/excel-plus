package com.ericyl.excel.writer.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;
import org.apache.poi.ss.util.CellRangeAddress;

@Data
@Accessors(chain = true)
@SuperBuilder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelRegion {

    private CellRangeAddress region;

    private ExcelColumn excelColumn;

}
