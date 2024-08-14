package com.ericyl.excel.example;

import com.ericyl.excel.writer.annotation.ExcelWriter;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class WriterObj {

    @ExcelWriter(value = "F5", height = 100)
    private String name;

    @ExcelWriter(value = "F6", align = "center", height = 1000)
    private Double money;

    @ExcelWriter(value = "F7", formatter = Writer1DateExcelWriterFormatter.class, width = 1000)
    private Date date;

    public WriterObj(String name, Double money) {
        this.name = name;
        this.money = money;
    }
}
