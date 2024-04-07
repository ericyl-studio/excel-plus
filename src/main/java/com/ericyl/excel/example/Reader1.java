package com.ericyl.excel.example;

import com.ericyl.excel.reader.annotation.ExcelReader;
import lombok.Data;

@Data
public class Reader1 {

    @ExcelReader(value = "A1")
    private String a1;

    @ExcelReader(value = "B5")
    private Double b5;


}
