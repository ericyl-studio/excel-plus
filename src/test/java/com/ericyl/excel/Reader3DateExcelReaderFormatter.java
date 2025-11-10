package com.ericyl.excel;

import com.ericyl.excel.reader.formatter.DateExcelReaderFormatter;

public class Reader3DateExcelReaderFormatter extends DateExcelReaderFormatter {
    @Override
    public String formatter(String str) {
        return "yyyy-MM-dd'T'HH:mm:ss";
    }
}
