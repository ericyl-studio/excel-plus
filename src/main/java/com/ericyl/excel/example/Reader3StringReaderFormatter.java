package com.ericyl.excel.example;

import com.ericyl.excel.reader.formatter.IExcelReaderFormatter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

public class Reader3StringReaderFormatter implements IExcelReaderFormatter<String> {
    @Override
    public String format(Cell cell) {
        String value = cell.getStringCellValue();
        if (StringUtils.isEmpty(value))
            return "未知";
        return value;
    }
}
