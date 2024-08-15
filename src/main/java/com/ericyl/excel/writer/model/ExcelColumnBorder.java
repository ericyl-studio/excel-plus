package com.ericyl.excel.writer.model;

import com.ericyl.excel.writer.common.BorderValue;
import lombok.Data;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

@Data
public class ExcelColumnBorder {

    BorderValue[] value = {};

    BorderStyle style = BorderStyle.NONE;

    IndexedColors color;

}
