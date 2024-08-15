package com.ericyl.excel.example;

import com.ericyl.excel.ExcelReaderUtils;
import com.ericyl.excel.ExcelWriterUtils;
import com.ericyl.excel.reader.IExcelReaderListener;
import com.ericyl.excel.writer.model.ExcelColumn;
import com.ericyl.excel.writer.model.ExcelTable;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.StreamSupport;

public class Test {

    public static void main(String... args) {
        //Reader
        try (InputStream inputStream = Files.newInputStream(new File("./demo.xlsx").toPath())) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Iterable<Sheet> sheetIterable = workbook::sheetIterator;
            StreamSupport.stream(sheetIterable.spliterator(), false).forEach(sheet -> {
                if (Objects.equals("Sheet1", sheet.getSheetName()))
                    read1(sheet);
                if (Objects.equals("Sheet2", sheet.getSheetName()))
                    read2(sheet);
                if (Objects.equals("Sheet3", sheet.getSheetName()))
                    read3(sheet);
            });
        } catch (Exception e) {
            e.printStackTrace();
        }
        //Writer
        writeObj();
        write1();
        write2();
        write3();

    }

    //单类读取
    private static void read1(Sheet sheet) {
        Reader1 reader1 = ExcelReaderUtils.doIt(sheet, Reader1.class);
        System.out.println(reader1);
    }

    //下标读取列表
    private static void read2(Sheet sheet) {
        List<Reader2> reader2List = ExcelReaderUtils.doList(sheet, Reader2.class, new IExcelReaderListener() {
            @Override
            public int headerNumber(Sheet sheet) {
                return IntStream.range(0, sheet.getLastRowNum()).filter(rowIndex -> {
                    Row row = sheet.getRow(rowIndex);
                    Iterable<Cell> cellIterable = row::cellIterator;
                    return StreamSupport.stream(cellIterable.spliterator(), false).anyMatch(cell -> {
                        Object obj = ExcelReaderUtils.getCellValue(cell);
                        return Objects.equals("名称", obj);
                    });
                }).findFirst().orElse(-1) + 1;


            }

            @Override
            public boolean isFooter(Row row) {
                Iterable<Cell> cellIterable = row::cellIterator;
                return StreamSupport.stream(cellIterable.spliterator(), false).anyMatch(cell -> {
                    Object obj = ExcelReaderUtils.getCellValue(cell);
                    return obj instanceof String && ((String) obj).contains("注：");
                });

            }
        });
        reader2List.forEach(System.out::println);
    }

    //表头读取列表
    private static void read3(Sheet sheet) {
        List<Reader3> reader3List = ExcelReaderUtils.doList(sheet, Reader3.class, new IExcelReaderListener() {

            @Override
            public int headerNumber(Sheet sheet) {
                return 3;
            }

            @Override
            public boolean isFooter(Row row) {
                return false;
            }
        });
        reader3List.forEach(System.out::println);
    }


    //数据写入
    private static void writeObj() {
        try (InputStream inputStream = Files.newInputStream(new File("./demo.xlsx").toPath())) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Iterable<Sheet> sheetIterable = workbook::sheetIterator;
            StreamSupport.stream(sheetIterable.spliterator(), false).forEach(sheet -> {
                if (Objects.equals("Sheet1", sheet.getSheetName())) {
                    WriterObj writerObj = new WriterObj("name_", 10.0, new Date());
                    ExcelWriterUtils.obj2Excel(workbook, sheet, writerObj);
                }
            });
            ExcelWriterUtils.toFile(workbook);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //列表生成
    private static void write1() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        List<Writer1> writer1List = IntStream.range(0, 10).mapToObj(index -> new Writer1("name_" + index, 10.0 + index, new Date())).collect(Collectors.toList());
        ExcelWriterUtils.list2Excel(workbook, sheet, writer1List, Writer1.class);
        ExcelWriterUtils.toFile(workbook);
    }

    //分页生成
    private static void write2() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        ExcelWriterUtils.list2Excel(workbook, sheet, 5, 10, Writer1.class,
                (pageNumber, pageSize) -> IntStream.range(0, pageSize).mapToObj(index -> new Writer1("name_" + pageNumber + "_" + index, 10.0 + index)).collect(Collectors.toList())
        );
        ExcelWriterUtils.toFile(workbook);
    }

    //复杂表结构生成
    private static void write3() {
        List<List<ExcelColumn>> headerList = new ArrayList<List<ExcelColumn>>() {{
            add(new ArrayList<ExcelColumn>() {{
                add(new ExcelColumn("项目", "xm", 4, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn("支出", "zc", 3, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
            }});
            add(new ArrayList<ExcelColumn>() {{
                add(new ExcelColumn("科目编码", "kmbm", 3, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn("科目名称", "kmmc", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn("小计", "xj", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn("基本支出", "jbzc", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn("项目支出", "mxzc", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
            }});
            add(new ArrayList<ExcelColumn>() {{
                add(new ExcelColumn("类", "l", 1, 2).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn("款", "k", 1, 2).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn("项", "x", 1, 2).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn("栏次", "lc", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn("1", "lc1", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn("2", "lc2", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn("3", "lc3", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
            }});
            add(new ArrayList<ExcelColumn>() {{
                add(new ExcelColumn("合计", "hj", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn(0.0, "hj1", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn(0.1, "hj2", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn(0.2, "hj3", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
            }});

        }};
        //可通过 List.stream 进行转换
        List<List<ExcelColumn>> bodyList = new ArrayList<List<ExcelColumn>>() {{
            add(new ArrayList<ExcelColumn>() {{
                add(new ExcelColumn("科目代码1", "kmdm", 3, 1).withHorizontalAlignment(HorizontalAlignment.LEFT));
                add(new ExcelColumn("科目名称1", "kmmc", 1, 1).withHorizontalAlignment(HorizontalAlignment.LEFT));
                add(new ExcelColumn(1.0, "xj", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn(2.0, "jbzc", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
                add(new ExcelColumn(3.0, "xmzc", 1, 1).withHorizontalAlignment(HorizontalAlignment.CENTER));
            }});
        }};
        List<List<ExcelColumn>> footerList = new ArrayList<List<ExcelColumn>>() {{
            add(new ArrayList<ExcelColumn>() {{
                add(new ExcelColumn("注：", ""));
            }});
        }};
        ExcelTable table = new ExcelTable(headerList, bodyList, footerList);

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        ExcelWriterUtils.table2Excel(workbook, sheet, table);
        ExcelWriterUtils.toFile(workbook);
    }

}
