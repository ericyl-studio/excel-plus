package com.ericyl.excel.writer;

/**
 * Excel写入监听器接口
 * <p>
 * 用于分页写入Excel时的数据获取。
 * 在处理大数据量导出时，可以通过实现此接口来分批获取数据，
 * 避免一次性加载所有数据导致的内存溢出问题。
 * </p>
 * 
 * @param <T> 数据类型
 * @author ericyl
 * @since 1.0
 */
public interface IExcelWriterListener<T> {

    /**
     * 分页获取数据
     * <p>
     * 根据页码和每页大小获取对应的数据。
     * 实现此方法时应该返回指定页的数据。
     * </p>
     * 
     * @param pageNumber 页码（从1开始）
     * @param pageSize   每页大小
     * @return 当前页的数据
     */
    T doSomething(int pageNumber, int pageSize);

}
