package com.ericyl.excel.writer.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;
import lombok.experimental.SuperBuilder;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.Objects;

@Data
@Accessors(chain = true)
@SuperBuilder
@NoArgsConstructor
@AllArgsConstructor
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

    @Override
    public int compareTo(ExcelColumn o) {
        return Objects.compare(getCellIndex(), o.getCellIndex(), Integer::compareTo);
    }
}
