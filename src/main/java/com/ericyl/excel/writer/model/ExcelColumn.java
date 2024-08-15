package com.ericyl.excel.writer.model;

import lombok.Data;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

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
    private VerticalAlignment verticalAlignment;
    private HorizontalAlignment horizontalAlignment;
    private ExcelColumnBorder border;

    public ExcelColumn(String key) {
        this(null, key, 1, 1);
    }

    public ExcelColumn(Object data, String key) {
        this(data, key, 1, 1);
    }

    public ExcelColumn(Object data, String key, int colspan, int rowspan) {
        this.data = data;
        this.colspan = colspan;
        this.rowspan = rowspan;
        this.verticalAlignment = VerticalAlignment.CENTER;
        this.key = key;
    }

    public ExcelColumn withHorizontalAlignment(HorizontalAlignment alignment) {
        this.horizontalAlignment = alignment;
        return this;
    }

    public ExcelColumn withVerticalAlignment(VerticalAlignment alignment) {
        this.verticalAlignment = alignment;
        return this;
    }

    public ExcelColumn withBorder(ExcelColumnBorder border) {
        this.border = border;
        return this;
    }

    @Override
    public int compareTo(ExcelColumn o) {
        return Objects.compare(getCellIndex(), o.getCellIndex(), Integer::compareTo);
    }
}
