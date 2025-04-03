package com.ericyl.excel.reader.model;

import com.ericyl.excel.reader.formatter.IExcelReaderFormatter;
import lombok.Data;

import java.lang.reflect.Field;

/**
 * 基础单元格
 */
@Data
public class FieldCell {
    /**
     * 属性
     */
    private Field field;
    /**
     * 行下标
     * 在列表时为null
     */
    private Integer rowIndex;
    /**
     * 列下标
     */
    private Integer startCellIndex;
    /**
     * 列下标
     */
    private Integer endCellIndex;
    /**
     * 数据转换器
     */
    private IExcelReaderFormatter<?> formatter;
}
