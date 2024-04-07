package com.ericyl.excel.writer.formatter;

public class DefaultExcelWriterFormatter implements IExcelWriterFormatter {
    @Override
    public Object format(Object o) {
        return o;
    }
}
