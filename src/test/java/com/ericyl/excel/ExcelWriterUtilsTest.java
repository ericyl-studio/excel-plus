package com.ericyl.excel;

import com.ericyl.excel.writer.model.ExcelColumn;
import com.ericyl.excel.writer.model.ExcelTable;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Arrays;
import java.util.Collections;

import static org.junit.jupiter.api.Assertions.assertEquals;

class ExcelWriterUtilsTest {

    @org.junit.jupiter.api.Test
    void table2ExcelSetsRowHeightFromColumns() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        ExcelTable table = new ExcelTable(
                Collections.singletonList(Arrays.asList(
                        new ExcelColumn("header one", "h1").setHeight(24.5F),
                        new ExcelColumn("header two", "h2").setHeight(32F))),
                Collections.singletonList(Arrays.asList(
                        new ExcelColumn("body one", "b1").setHeight(18F),
                        new ExcelColumn("body two", "b2").setHeight(28.25F))),
                Collections.singletonList(Collections.singletonList(
                        new ExcelColumn("footer", "f1").setHeight(21F))));

        ExcelWriterUtils.table2Excel(workbook, sheet, table);

        assertEquals(32F, sheet.getRow(0).getHeightInPoints(), 0.01F);
        assertEquals(28.25F, sheet.getRow(1).getHeightInPoints(), 0.01F);
        assertEquals(21F, sheet.getRow(2).getHeightInPoints(), 0.01F);
    }

    @org.junit.jupiter.api.Test
    void table2ExcelSetsColumnWidthFromColumns() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        ExcelTable table = new ExcelTable(
                Collections.singletonList(Arrays.asList(
                        new ExcelColumn("header one", "h1").setWidth(4096),
                        new ExcelColumn("header two", "h2").setWidth(6144))),
                Collections.singletonList(Arrays.asList(
                        new ExcelColumn("body one", "b1"),
                        new ExcelColumn("body two", "b2"))),
                Collections.emptyList());

        ExcelWriterUtils.table2Excel(workbook, sheet, table);

        assertEquals(4096, sheet.getColumnWidth(0));
        assertEquals(6144, sheet.getColumnWidth(1));
    }
}
