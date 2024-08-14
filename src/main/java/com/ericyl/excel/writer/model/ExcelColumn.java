package com.ericyl.excel.writer.model;

import lombok.Data;

import java.util.Objects;

@Data
public class ExcelColumn implements Comparable<ExcelColumn> {
    private String key;
    private Object data;
    private Integer rowIndex;
    private Integer cellIndex;
    private int colspan;
    private int rowspan;
    private Integer width;
    private Short height;
    private String align;
//    private IExcelWriterFormatter formatter;

    public ExcelColumn(String key) {
        this(null, key, 1, 1, "start");
    }

    public ExcelColumn(Object data, String key) {
        this(data, key, 1, 1, "start");
    }

    public ExcelColumn(Object data, String key, int colspan, int rowspan, String align) {
        this.data = data;
        this.colspan = colspan;
        this.rowspan = rowspan;
        this.align = align;
        this.key = key;
    }

//    public ExcelColumn<T> withName(String name) {
//        this.name = name;
//        return this;
//    }
//
//    public ExcelColumn<T> withColspan(int colspan) {
//        this.colspan = colspan;
//        return this;
//    }
//
//
//    public ExcelColumn<T> withRowspan(int rowspan) {
//        this.rowspan = rowspan;
//        return this;
//    }
//
//    public ExcelColumn<T> withAlign(String align) {
//        this.align = align;
//        return this;
//    }
//
//
//    public ExcelColumn<T> withWidth(int width) {
//        this.width = width;
//        return this;
//    }
//
//    public ExcelColumn<T> withHeight(int height) {
//        this.height = height;
//        return this;
//    }

    @Override
    public int compareTo(ExcelColumn o) {
        return Objects.compare(getCellIndex(), o.getCellIndex(), Integer::compareTo);
    }
}
