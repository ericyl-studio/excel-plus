package com.ericyl.excel.writer.annotation;

import com.ericyl.excel.writer.formatter.DefaultExcelWriterFormatter;
import com.ericyl.excel.writer.formatter.IExcelWriterFormatter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;

/**
 * Excel写入配置注解
 * <p>
 * 用于配置字段在Excel中的写入方式，包括位置、样式、格式化等。
 * 支持两种定位方式：
 * <ul>
 * <li>坐标定位：通过value属性指定单元格坐标（如"A1"）</li>
 * <li>列表定位：通过index属性指定列索引，配合name属性设置表头名称</li>
 * </ul>
 * </p>
 * 
 * @author ericyl
 * @since 1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD })
@Inherited
public @interface ExcelWriter {

    /**
     * 单元格坐标
     * <p>
     * 用于对象写入模式，指定数据写入的具体位置。
     * 格式：列标+行号，如 "A1", "B2", "AA10" 等。
     * </p>
     * 
     * @return 单元格坐标，默认为空
     */
    String value() default "";

    /**
     * 列表头名称
     * <p>
     * 用于列表写入模式，指定该字段对应的表头名称。
     * 如果为空，则使用字段名作为表头。
     * </p>
     * 
     * @return 表头名称，默认为空
     */
    String name() default "";

    /**
     * 列索引（排序）
     * <p>
     * 用于列表写入模式，指定该字段在Excel中的列位置。
     * 值越小越靠前，从0开始。
     * </p>
     * 
     * @return 列索引，默认为-1（不指定）
     */
    int index() default -1;

    /**
     * 列宽度
     * <p>
     * 设置Excel列的宽度，单位为字符宽度的1/256。
     * 例如：width = 256 * 10 表示10个字符宽度。
     * </p>
     * 
     * @return 列宽度，默认为-1（使用Excel默认宽度）
     */
    int width() default -1;

    /**
     * 行高度
     * <p>
     * 设置Excel行的高度，单位为点（1/20磅）。
     * 例如：height = 20 * 15 表示15磅高度。
     * </p>
     * 
     * @return 行高度，默认为-1（使用Excel默认高度）
     */
    short height() default -1;

    /**
     * 垂直对齐方式
     * <p>
     * 设置单元格内容的垂直对齐方式。
     * </p>
     * 
     * @return 垂直对齐方式，默认为居中对齐
     */
    VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;

    /**
     * 水平对齐方式
     * <p>
     * 设置单元格内容的水平对齐方式。
     * </p>
     * 
     * @return 水平对齐方式，默认为左对齐
     */
    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.LEFT;

    /**
     * 数据格式化器
     * <p>
     * 指定自定义的数据格式化器，用于在写入Excel前对数据进行格式化处理。
     * 例如：日期格式化、数字格式化等。
     * </p>
     * 
     * @return 格式化器类，默认使用DefaultExcelWriterFormatter
     */
    Class<? extends IExcelWriterFormatter> formatter() default DefaultExcelWriterFormatter.class;

    /**
     * 单元格边框设置
     * <p>
     * 配置单元格的边框样式，包括边框位置、样式和颜色。
     * </p>
     * 
     * @return 边框配置，默认为空配置
     */
    ExcelWriterBorder border() default @ExcelWriterBorder;
}
