package com.ericyl.excel.reader;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public interface IExcelReaderListener {

    //不是下标，从 1 开始
    default int startHeaderNumber(Sheet sheet) {
        return 1;
    }

    int headerNumber(Sheet sheet);

    boolean isFooter(Row row);
}