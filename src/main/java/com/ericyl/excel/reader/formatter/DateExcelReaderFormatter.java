package com.ericyl.excel.reader.formatter;


import com.ericyl.excel.ExcelReaderUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 时间类型基础数据转换器
 */
public abstract class DateExcelReaderFormatter implements IExcelReaderFormatter<Date> {

    public abstract String formatter(String str);

    @Override
    public Date format(Cell cell) {
        Object obj = ExcelReaderUtils.getCellValue(cell);
        if (obj == null)
            return null;
        if (obj instanceof Date)
            return (Date) obj;
        String cellValue;
        if (obj instanceof String)
            cellValue = cell.getStringCellValue();
        else if (Number.class.isAssignableFrom(obj.getClass()))
            cellValue = String.valueOf(((Number)obj).longValue());
        else
            throw new RuntimeException("Not support type of " + obj.getClass());
        if (StringUtils.isEmpty(cellValue))
            return null;
        String formatter = formatter(cellValue);
        if (StringUtils.isEmpty(formatter))
            return null;
        try {
            return new SimpleDateFormat(formatter(cellValue)).parse(cellValue);
        } catch (ParseException ex) {
            throw new RuntimeException(ex);
        }

    }

}
