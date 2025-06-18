package com.ericyl.excel.reader.formatter;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 数据转换器接口
 * 
 * @param <T> 转换后的数据类型
 */
public interface IExcelReaderFormatter<T> {

    T format(Cell cell);
}
