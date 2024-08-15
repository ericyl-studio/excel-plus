package com.ericyl.excel.example;

import com.ericyl.excel.writer.annotation.ExcelWriter;
import com.ericyl.excel.writer.annotation.ExcelWriterBorder;
import com.ericyl.excel.writer.common.BorderValue;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class WriterObj {

    @ExcelWriter(value = "F5", height = 100, border = @ExcelWriterBorder(value = {BorderValue.TOP}))
    private String name;

    @ExcelWriter(value = "F6", horizontalAlignment = HorizontalAlignment.CENTER, height = 1000)
    private Double money;

    @ExcelWriter(value = "F7", formatter = Writer1DateExcelWriterFormatter.class, width = 1000)
    private Date date;

    public WriterObj(String name, Double money) {
        this.name = name;
        this.money = money;
    }
}
