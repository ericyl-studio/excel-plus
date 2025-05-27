package com.ericyl.excel.reader;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Excel读取监听器接口
 * <p>
 * 用于自定义Excel读取过程中的行为，如确定表头位置、判断表尾等。
 * 实现此接口可以灵活处理不同格式的Excel文件。
 * </p>
 * 
 * @author ericyl
 * @since 1.0
 */
public interface IExcelReaderListener {

    /**
     * 获取表头开始行号
     * <p>
     * 用于处理多行表头的情况，返回表头的第一行行号。
     * 注意：返回的是实际行号（从1开始），不是索引（从0开始）。
     * </p>
     * 
     * @param sheet Excel工作表
     * @return 表头开始行号，默认为1（第一行）
     */
    default int startHeaderNumber(Sheet sheet) {
        return 1;
    }

    /**
     * 获取表头结束行号
     * <p>
     * 返回表头的最后一行行号。
     * 注意：返回的是实际行号（从1开始），不是索引（从0开始）。
     * </p>
     * 
     * @param sheet Excel工作表
     * @return 数据结束行号
     */
    int endHeaderNumber(Sheet sheet);

    /**
     * 判断是否为表尾行
     * <p>
     * 用于识别Excel中的表尾行（如合计行、备注行等），
     * 这些行通常不应该被当作数据行处理。
     * </p>
     * 
     * @param row 要判断的行
     * @return 如果是表尾行返回true，否则返回false
     */
    boolean isFooter(Row row);
}