package com.ericyl.excel.example;

import com.ericyl.excel.reader.annotation.ExcelReader;
import lombok.Data;

import java.util.Date;

@Data
public class Reader3 {

    @ExcelReader(name = "名称", formatter = Reader3StringReaderFormatter.class)
    private String t;

    @ExcelReader(name = {"合计111", "合计"})
    private Double t00;

    @ExcelReader(name = "合计")
    private Double t0;

    @ExcelReader(name = {"支出", "合计"})
    private Double t1;

    @ExcelReader(name = {"收入", "合计123"})
    private Double t2;

    @ExcelReader(name = "统计时间", formatter = Reader3DateExcelReaderFormatter.class)
    private Date t3;

    @ExcelReader(name = "创建时间")
    private Date t4;

}
