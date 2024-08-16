package com.ericyl.excel.writer.model;

import com.ericyl.excel.writer.common.BorderValue;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ExcelColumnBorder {

    BorderValue[] value = {};

    BorderStyle style = BorderStyle.NONE;

    IndexedColors color;

}
