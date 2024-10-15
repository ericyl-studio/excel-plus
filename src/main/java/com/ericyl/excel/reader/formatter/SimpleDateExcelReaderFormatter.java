package com.ericyl.excel.reader.formatter;

import org.apache.commons.lang3.StringUtils;

/**
 * 时间类型数据转换器
 */
public class SimpleDateExcelReaderFormatter extends DateExcelReaderFormatter {

    @Override
    public String formatter(String str) {
        if (StringUtils.isEmpty(str))
            return null;
        if (str.matches("^(\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2})$"))
            return "yyyy-MM-dd HH:mm:ss";
        if (str.matches("^(\\d{8})$"))
            return "yyyyMMdd";
        if (str.matches("^(\\d{4}-\\d{2}-\\d{2})$"))
            return "yyyy-MM-dd";
        if (str.matches("^(\\d{2}:\\d{2}:\\d{2})$"))
            return "HH:mm:ss";
        if (str.matches("^(\\d{4}年\\d{2}月\\d{2}日)$"))
            return "yyyy年MM月dd日";
        if (str.matches("^(\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2})$"))
            return "yyyy-MM-dd'T'HH:mm:ss";
        if (str.matches("^(\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2}.\\d{3})$"))
            return "yyyy-MM-dd'T'HH:mm:ss.SSS";
        if (str.matches("^(\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2}Z)$"))
            return "yyyy-MM-dd'T'HH:mm:ssZ";
        return null;
    }

}
