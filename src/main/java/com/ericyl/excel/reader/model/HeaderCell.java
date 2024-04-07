package com.ericyl.excel.reader.model;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class HeaderCell {

    private Object cellValue;
    private int rowIndex;
    private int startCellIndex;
    private int endCellIndex;

}
