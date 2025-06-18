package com.ericyl.excel.reader.annotation;

import com.ericyl.excel.reader.formatter.DefaultExcelReaderFormatter;
import com.ericyl.excel.reader.formatter.IExcelReaderFormatter;

import java.lang.annotation.*;

@Retention(RetentionPolicy.RUNTIME) // 使注解在运行时可用
@Target({ ElementType.FIELD })
@Inherited
public @interface ExcelReader {

    /**
     * 单元格坐标
     * 例如: "A1", "B1"
     * 
     * @return 单元格坐标字符串，默认为空
     */
    String value() default "";

    /**
     * 下标
     * 例如: 0, 1, 2
     * 
     * @return 列索引，默认为-1（不指定）
     */
    int index() default -1;

    /**
     * 名称
     * 支持多表头
     * 例如: {"支出", "合计"}, "名称"
     * 
     * @return 表头名称数组，默认为空数组
     */
    String[] name() default {};

    /**
     * 数据类型转换器
     * 
     * @return 数据转换器类，默认使用DefaultExcelReaderFormatter
     */
    Class<? extends IExcelReaderFormatter<?>> formatter() default DefaultExcelReaderFormatter.class;

}
