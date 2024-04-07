package com.ericyl.excel.reader;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public interface IExcelReaderListener {
    int headerNumber(Sheet sheet);

    boolean isFooter(Row row);
}