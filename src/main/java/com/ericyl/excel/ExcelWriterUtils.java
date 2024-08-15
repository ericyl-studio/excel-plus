package com.ericyl.excel;

import com.ericyl.excel.util.ObjectUtils;
import com.ericyl.excel.writer.IExcelWriterListener;
import com.ericyl.excel.writer.annotation.ExcelWriter;
import com.ericyl.excel.writer.formatter.DefaultExcelWriterFormatter;
import com.ericyl.excel.writer.formatter.IExcelWriterFormatter;
import com.ericyl.excel.writer.model.ExcelColumn;
import com.ericyl.excel.writer.model.ExcelTable;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

//导出 excel
public class ExcelWriterUtils {

    public static <T> void obj2Excel(Workbook workbook, Sheet sheet, T obj) {
        if (sheet == null)
            throw new RuntimeException("表格数据不能为空");

        List<ExcelColumn> excelColumnList = getExcelColumns(obj.getClass(), obj);

        excelColumnList.forEach(excelColumn -> {
            Row row = sheet.getRow(excelColumn.getRowIndex());
            if (row == null)
                row = sheet.createRow(excelColumn.getRowIndex());
            Cell cell = row.getCell(excelColumn.getCellIndex());
            if (cell == null)
                cell = row.createCell(excelColumn.getCellIndex());
            setCellValue(workbook, cell, excelColumn);
        });

    }

    public static <T> void list2Excel(Workbook workbook, Sheet sheet, List<T> list, Class<T> clazz) {
        if (CollectionUtils.isEmpty(list)) throw new RuntimeException("未查询到需导出的数据");
        // 表头
        Row title = sheet.createRow(0);
        List<ExcelColumn> titleExcelColumnList = getExcelColumns(clazz, null);
        IntStream.range(0, titleExcelColumnList.size()).forEach(index -> {
            ExcelColumn excelColumn = titleExcelColumnList.get(index);
            setCellWidth(sheet, index, excelColumn.getWidth());
            Cell cell = title.createCell(index);
            if (excelColumn.getData() == null) {
                cell.setCellValue(excelColumn.getKey());
                setCellStyle(workbook, cell, excelColumn);
            } else
                setCellValue(workbook, cell, excelColumn);

        });

        //内容
        IntStream.range(0, list.size()).forEach(index -> {
            Row row = sheet.createRow(index + 1);
            List<ExcelColumn> excelColumnList = getExcelColumns(clazz, list.get(index));

            Short height = excelColumnList.stream().map(ExcelColumn::getHeight).filter(Objects::nonNull).max(Comparator.naturalOrder()).orElse(null);
            setCellHeight(row, height);

            IntStream.range(0, titleExcelColumnList.size()).forEach(i -> {
                Cell cell = row.createCell(i);
                ExcelColumn titleExcelColumn = titleExcelColumnList.get(i);
                ExcelColumn excelColumn = excelColumnList.stream().filter(it -> Objects.equals(titleExcelColumn.getKey(), it.getKey())).findFirst().orElse(null);
                setCellValue(workbook, cell, excelColumn);
            });
        });


    }

    public static <T> void list2Excel(Workbook workbook, Sheet sheet, int page, int pageSize, Class<T> clazz, IExcelWriterListener<List<T>> doExcel) {

        // 表头
        Row title = sheet.createRow(0);
        List<ExcelColumn> titleExcelColumnList = getExcelColumns(clazz, null);
        IntStream.range(0, titleExcelColumnList.size()).forEach(index -> {
            ExcelColumn excelColumn = titleExcelColumnList.get(index);
            setCellWidth(sheet, index, excelColumn.getWidth());
            Cell cell = title.createCell(index);
            if (excelColumn.getData() == null) {
                cell.setCellValue(excelColumn.getKey());
                setCellStyle(workbook, cell, excelColumn);
            } else
                setCellValue(workbook, cell, excelColumn);
        });

        IntStream.range(1, page + 1).forEach(pageNumber -> {
            List<T> list = doExcel.doSomething(pageNumber, pageSize);

            IntStream.range(0, list.size()).forEach(index -> {
                Row row = sheet.createRow((pageNumber - 1) * pageSize + index + 1);
                List<ExcelColumn> excelColumnList = getExcelColumns(clazz, list.get(index));

                Short height = excelColumnList.stream().map(ExcelColumn::getHeight).filter(Objects::nonNull).max(Comparator.naturalOrder()).orElse(null);
                setCellHeight(row, height);

                IntStream.range(0, titleExcelColumnList.size()).forEach(i -> {
                    Cell cell = row.createCell(i);
                    ExcelColumn titleExcelColumn = titleExcelColumnList.get(i);
                    ExcelColumn excelColumn = excelColumnList.stream().filter(it -> Objects.equals(titleExcelColumn.getKey(), it.getKey())).findFirst().orElse(null);
                    setCellValue(workbook, cell, excelColumn);
                });
            });
        });

    }
    
    public static void table2Excel(Workbook workbook, Sheet sheet, ExcelTable table) {
        if (table == null) throw new RuntimeException("未查询到需导出的数据");
        setCell(workbook, sheet, table.getHeaders(), 0);
        setCell(workbook, sheet, table.getColumns(), table.getHeaders().size());
        setCell(workbook, sheet, table.getFooters(), table.getHeaders().size() + table.getColumns().size());
    }

    private static void setCell(Workbook workbook, Sheet sheet, List<List<ExcelColumn>> excelColumnList, int rowspan) {
        IntStream.range(0, excelColumnList.size()).forEach(index -> {
            Row row = sheet.createRow(rowspan + index);
            int parentColspan = IntStream.range(0, index).reduce(0, (acc, item) -> {
                List<ExcelColumn> excelColumns = excelColumnList.get(item);
//                acc += IntStream.range(0, excelColumns.size()).reduce(0, (acc1, item1) -> {
//                    ExcelColumn column = excelColumns.get(item1);
//                    acc1 += (column.getRowspan() - 1) > 0 ? column.getColspan() : 0;
//                    return acc1;
//                });
                acc += IntStream.range(0, excelColumns.size()).reduce(0, (acc1, item1) -> acc1 + ((excelColumns.get(item1).getRowspan() - 1) > 0 ? excelColumns.get(item1).getColspan() : 0));
                return acc;
            });
            List<ExcelColumn> columnList = excelColumnList.get(index);
            IntStream.range(0, columnList.size()).forEach(i -> {
                ExcelColumn excelColumn = columnList.get(i);
                int colspan = IntStream.range(0, i).reduce(0, (acc, item) -> acc + columnList.get(item).getColspan());
                Cell cell = row.createCell(colspan + parentColspan);
                setCellValue(workbook, cell, excelColumn);

                if (excelColumn.getColspan() > 1 || excelColumn.getRowspan() > 1) {
                    CellRangeAddress region = new CellRangeAddress(rowspan + index, rowspan + index + excelColumn.getRowspan() - 1, parentColspan + colspan, parentColspan + colspan + excelColumn.getColspan() - 1);
                    sheet.addMergedRegion(region);
                }
            });

        });
    }

    private static List<ExcelColumn> getExcelColumns(Class<?> clazz, Object obj) {
        Field[] fields = clazz.getDeclaredFields();
        return Arrays.stream(fields).map(field -> {
            ExcelColumn excelColumn = new ExcelColumn(field.getName());
            if (!field.isAnnotationPresent(ExcelWriter.class)) return excelColumn;
            ExcelWriter annotation = field.getAnnotation(ExcelWriter.class);
            if (obj == null) {
                //表头
                String cellName = annotation.name();
                if (StringUtils.isNotEmpty(cellName))
                    excelColumn.setData(cellName);
            } else {
                //内容
                Object data = ObjectUtils.getField(obj, field);
                if (data != null && annotation.formatter() != DefaultExcelWriterFormatter.class) {
                    try {
                        IExcelWriterFormatter writerFormatter = annotation.formatter().newInstance();
                        excelColumn.setData(writerFormatter.format(data));
                    } catch (Exception e) {
                        throw new RuntimeException(e);
                    }
                } else {
                    excelColumn.setData(data);
                }
            }

            String cellValue = annotation.value();
            int cellIndex = annotation.index();
            if (StringUtils.isNotEmpty(cellValue)) {
                //判断坐标
                Matcher matcher = Pattern.compile("(\\D+)(\\d+)").matcher(cellValue);
                if (matcher.find()) {
                    String[] parts = {matcher.group(1), matcher.group(2)};
                    excelColumn.setRowIndex(Integer.parseInt(parts[1]) - 1);
                    excelColumn.setCellIndex(ObjectUtils.convertToNumber(parts[0]) - 1);
                }

            } else if (cellIndex != -1) {
                excelColumn.setCellIndex(cellIndex);
            }
            int cellWidth = annotation.width();
            if (cellWidth != -1) excelColumn.setWidth(cellWidth);
            short cellHeight = annotation.height();
            if (cellHeight != -1) excelColumn.setHeight(cellHeight);
            String cellAlign = annotation.align();
            if (StringUtils.isNotEmpty(cellAlign)) excelColumn.setAlign(cellAlign);
            return excelColumn;
        }).sorted().collect(Collectors.toList());
    }

//    private static List<ExcelColumn> getHeaderExcelColumns(Class<?> clazz) {
//        Field[] fields = clazz.getDeclaredFields();
//        return Arrays.stream(fields).map(field -> {
//            ExcelColumn excelColumn = new ExcelColumn(field.getName());
//            if (!field.isAnnotationPresent(ExcelWriter.class)) return excelColumn;
//            ExcelWriter annotation = field.getAnnotation(ExcelWriter.class);
//            String cellName = annotation.name();
//            if (StringUtils.isNotEmpty(cellName))
//                excelColumn.setData(cellName);
//            int cellIndex = annotation.index();
//            if (cellIndex != -1) excelColumn.setCellIndex(cellIndex);
//            int cellWidth = annotation.width();
//            if (cellWidth != -1) excelColumn.setWidth(cellWidth);
//            short cellHeight = annotation.height();
//            if (cellHeight != -1) excelColumn.setHeight(cellHeight);
//            String cellAlign = annotation.align();
//            if (StringUtils.isNotEmpty(cellAlign)) excelColumn.setAlign(cellAlign);
//            return excelColumn;
//        }).sorted().collect(Collectors.toList());
//    }

    private static void setCellValue(Workbook workbook, Cell cell, ExcelColumn excelColumn) {
        if (excelColumn == null) return;
        Object obj = excelColumn.getData();
        if (obj == null) return;
        if (Number.class.isAssignableFrom(obj.getClass())) {
            cell.setCellValue(new BigDecimal(String.valueOf(obj)).doubleValue());
        } else if (obj instanceof String) {
            cell.setCellValue(obj.toString());
        } else if (obj instanceof Date) {
            cell.setCellValue((Date) obj);
        } else if (obj instanceof Boolean) {
            cell.setCellValue((Boolean) obj);
        }

        setCellStyle(workbook, cell, excelColumn);

    }

    private static void setCellStyle(Workbook workbook, Cell cell, ExcelColumn excelColumn) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        switch (excelColumn.getAlign()) {
            case "center":
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                break;
            case "end":
                cellStyle.setAlignment(HorizontalAlignment.RIGHT);
                break;
            case "start":
            default:
                cellStyle.setAlignment(HorizontalAlignment.LEFT);
                break;
        }

        cell.setCellStyle(cellStyle);
    }

    private static void setCellWidth(Sheet sheet, int index, Integer width) {
        if (width == null || width <= 0)
            return;
        sheet.setColumnWidth(index, width);
    }

    private static void setCellHeight(Row row, Short height) {
        if (height == null || height <= 0)
            return;
        row.setHeight(height);
    }

    public static String toFile(Workbook workbook) {
        String file = String.format("%s_%d.xlsx", UUID.randomUUID(), System.currentTimeMillis());

        try (OutputStream out = getOutputStream("excel" + File.separator + file)) {
            workbook.write(out);
            out.flush();
            return String.format("/excel/%s", file);
        } catch (IOException e) {
            throw new RuntimeException(e.getMessage());
        }
    }

    private static FileOutputStream getOutputStream(String path) throws IOException {
        File file = new File(path);
        if (!file.exists()) {
            file.mkdirs();
            file.delete();
            file.createNewFile();
        }
        return new FileOutputStream(file, false);
    }


}
