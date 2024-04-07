package com.ericyl.excel.reader.formatter;


import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Objects;

/**
 * 时间类型基础数据转换器
 */
public abstract class DateExcelReaderFormatter implements IExcelReaderFormatter<Date> {

    public abstract String formatter(String str);

    @Override
    public Date format(Cell cell) {
        try {
            return cell.getDateCellValue();
        } catch (IllegalStateException e) {
            if (!Objects.equals(CellType.STRING, cell.getCellType()))
                throw e;
            String cellValue = cell.getStringCellValue();
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

}
