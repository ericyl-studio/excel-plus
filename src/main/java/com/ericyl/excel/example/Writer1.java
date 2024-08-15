package com.ericyl.excel.example;

import com.ericyl.excel.writer.annotation.ExcelWriter;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class Writer1 {

    @ExcelWriter(name = "名称", index = 0, height = 100)
    private String name;

    @ExcelWriter(name = "金额", index = 1, horizontalAlignment = HorizontalAlignment.CENTER, height = 1000)
    private Double money;

    @ExcelWriter(name = "时间", index = 2, formatter = Writer1DateExcelWriterFormatter.class, width = 1000)
    private Date date;

    public Writer1(String name, Double money) {
        this.name = name;
        this.money = money;
    }
}
