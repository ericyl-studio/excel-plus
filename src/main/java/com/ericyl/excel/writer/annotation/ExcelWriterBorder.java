package com.ericyl.excel.writer.annotation;

import com.ericyl.excel.writer.common.BorderValue;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.*;

@Retention(RetentionPolicy.RUNTIME) // 使注解在运行时可用
@Target({ElementType.FIELD})
@Inherited
public @interface ExcelWriterBorder {

    BorderValue[] value() default {};

    BorderStyle style() default BorderStyle.THIN;

    IndexedColors color() default IndexedColors.BLACK;

}
