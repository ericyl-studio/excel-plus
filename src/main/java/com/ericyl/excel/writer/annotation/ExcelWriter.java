package com.ericyl.excel.writer.annotation;


import com.ericyl.excel.writer.formatter.DefaultExcelWriterFormatter;
import com.ericyl.excel.writer.formatter.IExcelWriterFormatter;

import java.lang.annotation.*;

@Retention(RetentionPolicy.RUNTIME) // 使注解在运行时可用
@Target({ElementType.FIELD})
@Inherited
public @interface ExcelWriter {

    /**
     * 单元格坐标
     * 例如: "A1", "B1"
     */
    String value() default "";

    /**
     * 名称
     */
    String name() default "";

    /**
     * 排序
     */
    int index() default -1;

    /**
     * 宽度
     */
    int width() default -1;

    /**
     * 高度
     * 取 max
     */
    short height() default -1;

    /**
     * 对齐方式
     */
    String align() default "start";

    /**
     * 数据类型转换器
     */
    Class<? extends IExcelWriterFormatter> formatter() default DefaultExcelWriterFormatter.class;

}

