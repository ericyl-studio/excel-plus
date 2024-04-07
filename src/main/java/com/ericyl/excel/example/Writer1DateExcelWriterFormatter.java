package com.ericyl.excel.example;

import com.ericyl.excel.writer.formatter.DateExcelWriterFormatter;

public class Writer1DateExcelWriterFormatter extends DateExcelWriterFormatter {
    @Override
    public String formatter() {
        return "yyyy-MM-dd'T'HH:mm:ss";
    }
}
