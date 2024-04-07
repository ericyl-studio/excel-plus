package com.ericyl.excel.example;

import com.ericyl.excel.reader.annotation.ExcelReader;
import lombok.Data;

@Data
public class Reader2 {

    @ExcelReader(index = 0)
    private String t0;

    @ExcelReader(index = 1)
    private Double t1;

}
