package com.ericyl.excel.example;

import com.ericyl.excel.ExcelReaderUtils;
import com.ericyl.excel.ExcelWriterUtils;
import com.ericyl.excel.reader.IExcelReaderListener;
import com.ericyl.excel.reader.model.HeaderCell;
import com.ericyl.excel.writer.common.BorderValue;
import com.ericyl.excel.writer.model.ExcelColumn;
import com.ericyl.excel.writer.model.ExcelColumnBorder;
import com.ericyl.excel.writer.model.ExcelTable;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.StreamSupport;

/**
 * Excel Plus 使用示例
 * <p>
 * 演示各种Excel读写功能的使用方法
 * </p>
 * 
 * @author ericyl
 * @since 1.0
 */
public class Test {

    public static void main(String... args) {
        // Excel 读取示例
        try (InputStream inputStream = Files.newInputStream(new File("./demo.xlsx").toPath())) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Iterable<Sheet> sheetIterable = workbook::sheetIterator;
            StreamSupport.stream(sheetIterable.spliterator(), false).forEach(sheet -> {
                // 根据不同的Sheet演示不同的读取方式
                // if (Objects.equals("Sheet1", sheet.getSheetName()))
                // read1(sheet); // 单对象读取
                // if (Objects.equals("Sheet2", sheet.getSheetName()))
                // read2(sheet); // 通过索引读取列表
                if (Objects.equals("Sheet3", sheet.getSheetName())) {
                    // read3(sheet); // 通过表头读取列表
                    read4(sheet); // 读取为Map格式
                }
            });
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Excel 写入示例
        // writeObj(); // 对象写入
        // write1(); // 列表写入
        // write2(); // 分页写入
        // write3(); // 复杂表格写入
    }

    /**
     * 示例1：单对象读取
     * <p>
     * 通过 @ExcelReader(value = "坐标") 读取指定单元格的数据
     * </p>
     */
    private static void read1(Sheet sheet) {
        Reader1 reader1 = ExcelReaderUtils.doIt(sheet, Reader1.class);
        System.out.println(reader1);
    }

    /**
     * 示例2：通过索引读取列表
     * <p>
     * 通过 @ExcelReader(index = 索引) 读取列数据
     * 演示如何自定义表头识别和表尾判断
     * </p>
     */
    private static void read2(Sheet sheet) {
        List<Reader2> reader2List = ExcelReaderUtils.doList(sheet, Reader2.class, new IExcelReaderListener() {
            @Override
            public int headerNumber(Sheet sheet) {
                // 动态查找包含"名称"的行作为表头行
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
                // 判断包含"注："的行为表尾
                Iterable<Cell> cellIterable = row::cellIterator;
                return StreamSupport.stream(cellIterable.spliterator(), false).anyMatch(cell -> {
                    Object obj = ExcelReaderUtils.getCellValue(cell);
                    return obj instanceof String && ((String) obj).contains("注：");
                });
            }
        });
        reader2List.forEach(System.out::println);
    }

    /**
     * 示例3：通过表头名称读取列表
     * <p>
     * 通过 @ExcelReader(name = {"表头"}) 读取列数据
     * 支持多级表头的匹配
     * </p>
     */
    private static void read3(Sheet sheet) {
        List<Reader3> reader3List = ExcelReaderUtils.doList(sheet, Reader3.class, new IExcelReaderListener() {
            @Override
            public int headerNumber(Sheet sheet) {
                return 3; // 数据从第3行开始
            }

            @Override
            public boolean isFooter(Row row) {
                return false; // 没有表尾
            }
        });
        reader3List.forEach(System.out::println);
    }

    /**
     * 示例4：读取为Map格式
     * <p>
     * 不需要定义实体类，直接将数据读取为Map
     * 适用于动态表头或临时数据处理
     * </p>
     */
    private static void read4(Sheet sheet) {
        IExcelReaderListener listener = new IExcelReaderListener() {
            @Override
            public int startHeaderNumber(Sheet sheet) {
                return 2; // 表头从第2行开始
            }

            @Override
            public int headerNumber(Sheet sheet) {
                return 3; // 数据从第3行开始
            }

            @Override
            public boolean isFooter(Row row) {
                return false;
            }
        };
        // 获取表头信息
        List<HeaderCell> headerCellList = ExcelReaderUtils.getHeaders(sheet, true, listener);
        // 读取数据为Map列表
        List<Map<String, Object>> list = ExcelReaderUtils.doMap(sheet, headerCellList, listener);
        list.forEach(System.out::println);
    }

    /**
     * 示例5：对象数据写入
     * <p>
     * 通过 @ExcelWriter(value = "坐标") 将对象数据写入指定位置
     * </p>
     */
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
            ExcelWriterUtils.toFile("excel", workbook);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 示例6：列表数据生成Excel
     * <p>
     * 通过 @ExcelWriter 注解配置表头、样式等
     * </p>
     */
    private static void write1() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        // 生成测试数据
        List<Writer1> writer1List = IntStream.range(0, 10)
                .mapToObj(index -> new Writer1("name_" + index, 10.0 + index, new Date()))
                .collect(Collectors.toList());
        // 写入Excel
        ExcelWriterUtils.list2Excel(workbook, sheet, writer1List, Writer1.class);
        // 保存文件
        ExcelWriterUtils.toFile("excel", workbook);
    }

    /**
     * 示例7：分页数据生成Excel
     * <p>
     * 适用于大数据量导出，避免内存溢出
     * </p>
     */
    private static void write2() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        // 分页写入，5页，每页10条
        ExcelWriterUtils.list2Excel(workbook, sheet, 5, 10, Writer1.class,
                (pageNumber, pageSize) -> IntStream.range(0, pageSize)
                        .mapToObj(index -> new Writer1("name_" + pageNumber + "_" + index, 10.0 + index, new Date()))
                        .collect(Collectors.toList()));
        ExcelWriterUtils.toFile("excel", workbook);
    }

    /**
     * 示例8：复杂表格结构生成
     * <p>
     * 支持多级表头、合并单元格、自定义样式等
     * </p>
     */
    private static void write3() {
        // 构建多级表头
        List<List<ExcelColumn>> headerList = new ArrayList<List<ExcelColumn>>() {
            {
                // 第一行表头
                add(new ArrayList<ExcelColumn>() {
                    {
                        add(new ExcelColumn("项目", "xm", 4, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER)
                                .withBorder(new ExcelColumnBorder(new BorderValue[] { BorderValue.ALL },
                                        BorderStyle.THIN, IndexedColors.BLACK)));
                        add(new ExcelColumn("支出", "zc", 3, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                    }
                });
                // 第二行表头
                add(new ArrayList<ExcelColumn>() {
                    {
                        add(new ExcelColumn("科目编码", "kmbm", 3, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                        add(new ExcelColumn("科目名称", "kmmc", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                        add(new ExcelColumn("小计", "xj", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                        add(new ExcelColumn("基本支出", "jbzc", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                        add(new ExcelColumn("项目支出", "mxzc", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                    }
                });
                // 第三行表头（包含跨行）
                add(new ArrayList<ExcelColumn>() {
                    {
                        add(new ExcelColumn("类", "l", 1, 2) // 跨2行
                                .withHorizontalAlignment(HorizontalAlignment.CENTER)
                                .withBorder(new ExcelColumnBorder(new BorderValue[] { BorderValue.ALL },
                                        BorderStyle.THIN, IndexedColors.BLACK)));
                        add(new ExcelColumn("款", "k", 1, 2) // 跨2行
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                        add(new ExcelColumn("项", "x", 1, 2) // 跨2行
                                .withHorizontalAlignment(HorizontalAlignment.CENTER)
                                .withBorder(new ExcelColumnBorder(new BorderValue[] { BorderValue.ALL },
                                        BorderStyle.THIN, IndexedColors.BLACK)));
                        add(new ExcelColumn("栏次", "lc", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                        add(new ExcelColumn("1", "lc1", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                        add(new ExcelColumn("2", "lc2", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                        add(new ExcelColumn("3", "lc3", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                    }
                });
                // 第四行表头
                add(new ArrayList<ExcelColumn>() {
                    {
                        add(new ExcelColumn("合计", "hj", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                        add(new ExcelColumn(0.0, "hj1", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                        add(new ExcelColumn(0.1, "hj2", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                        add(new ExcelColumn(0.2, "hj3", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                    }
                });
            }
        };

        // 构建数据内容（可通过 List.stream 进行转换）
        List<List<ExcelColumn>> bodyList = new ArrayList<List<ExcelColumn>>() {
            {
                add(new ArrayList<ExcelColumn>() {
                    {
                        add(new ExcelColumn("科目代码1", "kmdm", 3, 1)
                                .withHorizontalAlignment(HorizontalAlignment.LEFT));
                        add(new ExcelColumn("科目名称1", "kmmc", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.LEFT));
                        add(new ExcelColumn(1.0, "xj", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                        add(new ExcelColumn(2.0, "jbzc", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                        add(new ExcelColumn(3.0, "xmzc", 1, 1)
                                .withHorizontalAlignment(HorizontalAlignment.CENTER));
                    }
                });
            }
        };

        // 构建表尾
        List<List<ExcelColumn>> footerList = new ArrayList<List<ExcelColumn>>() {
            {
                add(new ArrayList<ExcelColumn>() {
                    {
                        add(new ExcelColumn("注：本表为示例数据", ""));
                    }
                });
            }
        };

        // 创建表格结构
        ExcelTable table = new ExcelTable(headerList, bodyList, footerList);

        // 生成Excel
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        ExcelWriterUtils.table2Excel(workbook, sheet, table);
        ExcelWriterUtils.toFile("excel/test", workbook);
    }
}
