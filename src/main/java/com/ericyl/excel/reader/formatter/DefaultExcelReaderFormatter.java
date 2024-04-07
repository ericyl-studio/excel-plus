package com.ericyl.excel.reader.formatter;



import org.apache.poi.ss.usermodel.Cell;

/**
 * 数据转换器默认实现
 */
public class DefaultExcelReaderFormatter implements IExcelReaderFormatter<Object> {
    @Override
    public Object format(Cell cell) {
        return cell;
    }
}

