package com.ericyl.excel.writer;

public interface IExcelWriterListener<T> {

    T doSomething(int pageNumber, int pageSize);

}
