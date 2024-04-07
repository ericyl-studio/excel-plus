package com.ericyl.excel.writer.formatter;


import org.apache.commons.lang3.StringUtils;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 时间类型基础数据转换器
 */
public abstract class DateExcelWriterFormatter implements IExcelWriterFormatter {

    public abstract String formatter();

    @Override
    public Object format(Object obj) {
        if (obj == null)
            return null;
        if (!(obj instanceof Date))
            return obj;
        String formatter = formatter();
        if (StringUtils.isEmpty(formatter))
            formatter = "yyyy-MM-dd";
        return new SimpleDateFormat(formatter).format(obj);
    }
}
