package com.ericyl.excel;

import com.ericyl.excel.util.ObjectUtils;
import com.ericyl.excel.writer.IExcelWriterListener;
import com.ericyl.excel.writer.annotation.ExcelWriter;
import com.ericyl.excel.writer.annotation.ExcelWriterBorder;
import com.ericyl.excel.writer.common.BorderValue;
import com.ericyl.excel.writer.formatter.DefaultExcelWriterFormatter;
import com.ericyl.excel.writer.formatter.IExcelWriterFormatter;
import com.ericyl.excel.writer.model.ExcelColumn;
import com.ericyl.excel.writer.model.ExcelColumnBorder;
import com.ericyl.excel.writer.model.ExcelRegion;
import com.ericyl.excel.writer.model.ExcelTable;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

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

/**
 * Excel 写入工具类
 * <p>
 * 提供多种方式生成 Excel 文件：
 * 1. 坐标写入：指定单元格坐标写入数据
 * 2. 对象写入：将对象数据写入到指定位置
 * 3. 列表写入：将列表数据批量写入（支持分页）
 * 4. 表格写入：支持复杂表格结构（多表头、合并单元格等）
 * </p>
 * 
 * @author ericyl
 * @since 1.0
 */
public class ExcelWriterUtils {

    /**
     * 坐标方式写入数据
     * <p>
     * 根据坐标（如 "A1"）将数据写入到指定单元格
     * </p>
     * 
     * @param workbook Excel工作簿
     * @param sheet    工作表
     * @param xy       坐标位置（如 "A1", "B2"）
     * @param obj      要写入的数据
     * @throws RuntimeException 当参数无效时抛出
     */
    public static void xy(Workbook workbook, Sheet sheet, String xy, Object obj) {
        if (sheet == null)
            throw new RuntimeException("表格数据不能为空");

        if (StringUtils.isEmpty(xy))
            throw new RuntimeException("坐标不能为空");

        if (obj == null)
            throw new RuntimeException("数据不能为空");

        // 解析坐标
        Matcher matcher = Pattern.compile("(\\D+)(\\d+)").matcher(xy);
        if (!matcher.find())
            return;

        String[] parts = { matcher.group(1), matcher.group(2) };
        int rowIndex = Integer.parseInt(parts[1]) - 1;
        int cellIndex = ObjectUtils.convertToNumber(parts[0]) - 1;

        // 获取或创建行
        Row row = sheet.getRow(rowIndex);
        if (row == null)
            row = sheet.createRow(rowIndex);

        // 获取或创建单元格
        Cell cell = row.getCell(cellIndex);
        if (cell == null)
            cell = row.createCell(cellIndex);

        ExcelColumn excelColumn = new ExcelColumn(obj, null);
        setCellValue(workbook, cell, excelColumn);
    }

    /**
     * 对象方式写入数据
     * <p>
     * 根据对象字段上的 @ExcelWriter 注解配置，将对象数据写入到相应位置
     * </p>
     * 
     * @param workbook Excel工作簿
     * @param sheet    工作表
     * @param obj      要写入的对象
     * @param <T>      对象类型
     * @throws RuntimeException 当表格数据为空时抛出
     */
    public static <T> void obj2Excel(Workbook workbook, Sheet sheet, T obj) {
        if (sheet == null)
            throw new RuntimeException("表格数据不能为空");

        List<ExcelColumn> excelColumnList = getExcelColumns(obj.getClass(), obj);

        for (ExcelColumn excelColumn : excelColumnList) {
            if (excelColumn.getRowIndex() == -1 || excelColumn.getCellIndex() == -1)
                continue;

            // 获取或创建行
            Row row = sheet.getRow(excelColumn.getRowIndex());
            if (row == null)
                row = sheet.createRow(excelColumn.getRowIndex());

            // 获取或创建单元格
            Cell cell = row.getCell(excelColumn.getCellIndex());
            if (cell == null)
                cell = row.createCell(excelColumn.getCellIndex());

            setCellValue(workbook, cell, excelColumn);
        }
    }

    /**
     * 列表方式写入数据
     * <p>
     * 将列表数据写入Excel，自动生成表头，支持自定义列宽、行高、对齐方式等
     * </p>
     * 
     * @param workbook Excel工作簿
     * @param sheet    工作表
     * @param list     数据列表
     * @param clazz    列表元素类型
     * @param <T>      元素类型
     * @throws RuntimeException 当列表为空时抛出
     */
    public static <T> void list2Excel(Workbook workbook, Sheet sheet, List<T> list, Class<T> clazz) {
        if (CollectionUtils.isEmpty(list))
            throw new RuntimeException("未查询到需导出的数据");

        // 生成表头
        Row title = sheet.createRow(0);
        List<ExcelColumn> titleExcelColumnList = getExcelColumns(clazz, null);

        IntStream.range(0, titleExcelColumnList.size()).forEach(index -> {
            ExcelColumn excelColumn = titleExcelColumnList.get(index);
            // 设置列宽
            setCellWidth(sheet, index, excelColumn.getWidth());

            Cell cell = title.createCell(index);
            if (excelColumn.getData() == null) {
                // 设置表头名称
                cell.setCellValue(excelColumn.getKey());
                setCellStyle(workbook, cell, excelColumn);
            } else
                setCellValue(workbook, cell, excelColumn);
        });

        // 写入内容
        IntStream.range(0, list.size()).forEach(index -> {
            Row row = sheet.createRow(index + 1);
            List<ExcelColumn> excelColumnList = getExcelColumns(clazz, list.get(index));

            // 设置行高（取最大值）
            Short height = excelColumnList.stream()
                    .map(ExcelColumn::getHeight)
                    .filter(Objects::nonNull)
                    .max(Comparator.naturalOrder())
                    .orElse(null);
            setCellHeight(row, height);

            // 写入每个单元格
            IntStream.range(0, titleExcelColumnList.size()).forEach(i -> {
                Cell cell = row.createCell(i);
                ExcelColumn titleExcelColumn = titleExcelColumnList.get(i);
                ExcelColumn excelColumn = excelColumnList.stream()
                        .filter(it -> Objects.equals(titleExcelColumn.getKey(), it.getKey()))
                        .findFirst()
                        .orElse(null);
                setCellValue(workbook, cell, excelColumn);
            });
        });
    }

    /**
     * 分页方式写入数据
     * <p>
     * 支持大数据量分页写入，避免内存溢出
     * </p>
     * 
     * @param workbook Excel工作簿
     * @param sheet    工作表
     * @param page     总页数
     * @param pageSize 每页大小
     * @param clazz    列表元素类型
     * @param doExcel  数据获取监听器，用于分页获取数据
     * @param <T>      元素类型
     */
    public static <T> void list2Excel(Workbook workbook, Sheet sheet, int page, int pageSize,
            Class<T> clazz, IExcelWriterListener<List<T>> doExcel) {

        // 生成表头
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

        // 分页写入数据
        IntStream.range(1, page + 1).forEach(pageNumber -> {
            List<T> list = doExcel.doSomething(pageNumber, pageSize);

            IntStream.range(0, list.size()).forEach(index -> {
                Row row = sheet.createRow((pageNumber - 1) * pageSize + index + 1);
                List<ExcelColumn> excelColumnList = getExcelColumns(clazz, list.get(index));

                // 设置行高
                Short height = excelColumnList.stream()
                        .map(ExcelColumn::getHeight)
                        .filter(Objects::nonNull)
                        .max(Comparator.naturalOrder())
                        .orElse(null);
                setCellHeight(row, height);

                // 写入每个单元格
                IntStream.range(0, titleExcelColumnList.size()).forEach(i -> {
                    Cell cell = row.createCell(i);
                    ExcelColumn titleExcelColumn = titleExcelColumnList.get(i);
                    ExcelColumn excelColumn = excelColumnList.stream()
                            .filter(it -> Objects.equals(titleExcelColumn.getKey(), it.getKey()))
                            .findFirst()
                            .orElse(null);
                    setCellValue(workbook, cell, excelColumn);
                });
            });
        });
    }

    /**
     * 复杂表格方式写入数据
     * <p>
     * 支持多表头、表尾、合并单元格等复杂表格结构
     * </p>
     * 
     * @param workbook Excel工作簿
     * @param sheet    工作表
     * @param table    表格结构定义
     * @throws RuntimeException 当表格数据为空时抛出
     */
    public static void table2Excel(Workbook workbook, Sheet sheet, ExcelTable table) {
        if (table == null)
            throw new RuntimeException("未查询到需导出的数据");

        // 写入表头
        setCell(workbook, sheet, table.getHeaders(), 0);
        // 写入内容
        setCell(workbook, sheet, table.getColumns(), table.getHeaders().size());
        // 写入表尾
        setCell(workbook, sheet, table.getFooters(), table.getHeaders().size() + table.getColumns().size());
    }

    /**
     * 设置单元格数据（支持合并单元格）
     * 
     * @param workbook        Excel工作簿
     * @param sheet           工作表
     * @param excelColumnList 单元格数据列表
     * @param rowspan         起始行偏移
     */
    private static void setCell(Workbook workbook, Sheet sheet, List<List<ExcelColumn>> excelColumnList, int rowspan) {
        List<ExcelRegion> excelRegionList = new ArrayList<>();

        IntStream.range(0, excelColumnList.size()).forEach(index -> {
            Row row = sheet.createRow(rowspan + index);

            // 计算前面行占用的列数（处理跨行的情况）
            int parentColspan = IntStream.range(0, index).reduce(0, (acc, item) -> {
                List<ExcelColumn> excelColumns = excelColumnList.get(item);
                acc += IntStream.range(0, excelColumns.size()).reduce(0,
                        (acc1, item1) -> acc1
                                + ((excelColumns.get(item1).getRowspan() - 1) > 0 ? excelColumns.get(item1).getColspan()
                                        : 0));
                return acc;
            });

            List<ExcelColumn> columnList = excelColumnList.get(index);
            IntStream.range(0, columnList.size()).forEach(i -> {
                ExcelColumn excelColumn = columnList.get(i);
                // 计算当前行前面列占用的宽度
                int colspan = IntStream.range(0, i).reduce(0,
                        (acc, item) -> acc + columnList.get(item).getColspan());

                Cell cell = row.createCell(colspan + parentColspan);
                setCellValue(workbook, cell, excelColumn);

                // 处理合并单元格
                if (excelColumn.getColspan() > 1 || excelColumn.getRowspan() > 1) {
                    CellRangeAddress region = new CellRangeAddress(
                            rowspan + index,
                            rowspan + index + excelColumn.getRowspan() - 1,
                            parentColspan + colspan,
                            parentColspan + colspan + excelColumn.getColspan() - 1);
                    sheet.addMergedRegion(region);
                    excelRegionList.add(new ExcelRegion(region, excelColumn));
                }
            });
        });

        // 设置合并单元格的样式
        excelRegionList.forEach(excelRegion -> setRegionStyle(workbook, sheet, excelRegion.getRegion(), excelRegion.getExcelColumn()));
    }

    /**
     * 设置合并单元格的样式
     * <p>
     * 合并单元格中的每个单元格都需要设置样式
     * </p>
     * 
     * @param workbook    Excel工作簿
     * @param sheet       工作表
     * @param region      合并区域
     * @param excelColumn 单元格配置
     */
    private static void setRegionStyle(Workbook workbook, Sheet sheet, CellRangeAddress region,
            ExcelColumn excelColumn) {
        for (int rowNum = region.getFirstRow(); rowNum <= region.getLastRow(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null)
                row = sheet.createRow(rowNum);

            for (int colNum = region.getFirstColumn(); colNum <= region.getLastColumn(); colNum++) {
                Cell cell = row.getCell(colNum);
                if (cell == null)
                    cell = row.createCell(colNum);
                setCellStyle(workbook, cell, excelColumn);
            }
        }
    }

    /**
     * 解析Excel列配置
     * <p>
     * 根据 @ExcelWriter 注解配置，解析字段对应的Excel列信息
     * </p>
     * 
     * @param clazz 类型
     * @param obj   对象实例（为null时解析表头）
     * @return Excel列配置列表
     */
    private static List<ExcelColumn> getExcelColumns(Class<?> clazz, Object obj) {
        Field[] fields = clazz.getDeclaredFields();
        return Arrays.stream(fields).map(field -> {
            ExcelColumn excelColumn = new ExcelColumn(field.getName());
            if (!field.isAnnotationPresent(ExcelWriter.class))
                return excelColumn;

            ExcelWriter annotation = field.getAnnotation(ExcelWriter.class);

            if (annotation == null)
                return excelColumn;

            if (obj == null) {
                // 解析表头
                String cellName = annotation.name();
                if (StringUtils.isNotEmpty(cellName))
                    excelColumn.setData(cellName);
            } else {
                // 解析内容
                Object data = ObjectUtils.getField(obj, field);
                if (data != null && annotation.formatter() != DefaultExcelWriterFormatter.class) {
                    try {
                        IExcelWriterFormatter writerFormatter = annotation.formatter().newInstance();
                        excelColumn.setData(writerFormatter.format(data));
                    } catch (Exception e) {
                        throw new RuntimeException("创建格式化器失败: " + e.getMessage(), e);
                    }
                } else {
                    excelColumn.setData(data);
                }
            }

            // 解析坐标或索引
            String cellValue = annotation.value();
            int cellIndex = annotation.index();
            if (StringUtils.isNotEmpty(cellValue)) {
                // 坐标方式
                Matcher matcher = Pattern.compile("(\\D+)(\\d+)").matcher(cellValue);
                if (matcher.find()) {
                    String[] parts = { matcher.group(1), matcher.group(2) };
                    excelColumn.setRowIndex(Integer.parseInt(parts[1]) - 1);
                    excelColumn.setCellIndex(ObjectUtils.convertToNumber(parts[0]) - 1);
                }
            } else if (cellIndex != -1) {
                // 索引方式
                excelColumn.setCellIndex(cellIndex);
            }

            // 设置其他属性
            int cellWidth = annotation.width();
            if (cellWidth != -1)
                excelColumn.setWidth(cellWidth);

            short cellHeight = annotation.height();
            if (cellHeight != -1)
                excelColumn.setHeight(cellHeight);

            VerticalAlignment verticalAlignment = annotation.verticalAlignment();
            if (verticalAlignment != null)
                excelColumn.setVerticalAlignment(verticalAlignment);

            HorizontalAlignment horizontalAlignment = annotation.horizontalAlignment();
            if (horizontalAlignment != null)
                excelColumn.setHorizontalAlignment(horizontalAlignment);

            // 设置边框
            ExcelWriterBorder border = annotation.border();
            if (border != null) {
                ExcelColumnBorder excelColumnBorder = new ExcelColumnBorder();
                BorderValue[] values = border.value();
                if (ArrayUtils.isNotEmpty(values))
                    excelColumnBorder.setValue(values);
                BorderStyle style = border.style();
                if (style != null)
                    excelColumnBorder.setStyle(style);
                IndexedColors color = border.color();
                if (color != null)
                    excelColumnBorder.setColor(color);
                excelColumn.setBorder(excelColumnBorder);
            }

            return excelColumn;
        }).sorted().collect(Collectors.toList());
    }

    /**
     * 设置单元格值
     * <p>
     * 根据数据类型设置单元格的值，并应用样式
     * </p>
     * 
     * @param workbook    Excel工作簿
     * @param cell        单元格
     * @param excelColumn 单元格配置
     */
    private static void setCellValue(Workbook workbook, Cell cell, ExcelColumn excelColumn) {
        if (excelColumn == null)
            return;

        Object obj = excelColumn.getData();
        if (obj == null)
            return;

        // 根据数据类型设置值
        if (Number.class.isAssignableFrom(obj.getClass())) {
            cell.setCellValue(new BigDecimal(String.valueOf(obj)).doubleValue());
        } else if (obj instanceof String) {
            cell.setCellValue(obj.toString());
        } else if (obj instanceof Date) {
            cell.setCellValue((Date) obj);
        } else if (obj instanceof Boolean) {
            cell.setCellValue((Boolean) obj);
        } else {
            throw new RuntimeException("暂不支持当前数据类型: " + obj.getClass().getName());
        }

        // 应用样式
        setCellStyle(workbook, cell, excelColumn);
    }

    /**
     * 设置单元格样式
     * <p>
     * 包括边框、对齐方式等样式设置
     * </p>
     * 
     * @param workbook    Excel工作簿
     * @param cell        单元格
     * @param excelColumn 单元格配置
     */
    private static void setCellStyle(Workbook workbook, Cell cell, ExcelColumn excelColumn) {
        CellStyle cellStyle = workbook.createCellStyle();

        // 设置边框样式
        ExcelColumnBorder excelColumnBorder = excelColumn.getBorder();
        if (excelColumnBorder != null && ArrayUtils.isNotEmpty(excelColumnBorder.getValue())) {
            BorderStyle style = excelColumnBorder.getStyle();
            IndexedColors color = excelColumnBorder.getColor();

            for (BorderValue value : excelColumnBorder.getValue()) {
                switch (value) {
                    case ALL:
                        // 设置所有边框
                        cellStyle.setBorderTop(style);
                        cellStyle.setBorderBottom(style);
                        cellStyle.setBorderLeft(style);
                        cellStyle.setBorderRight(style);
                        if (color != null) {
                            cellStyle.setTopBorderColor(color.getIndex());
                            cellStyle.setBottomBorderColor(color.getIndex());
                            cellStyle.setLeftBorderColor(color.getIndex());
                            cellStyle.setRightBorderColor(color.getIndex());
                        }
                        break;
                    case X:
                        // 设置左右边框
                        cellStyle.setBorderLeft(style);
                        cellStyle.setBorderRight(style);
                        if (color != null) {
                            cellStyle.setLeftBorderColor(color.getIndex());
                            cellStyle.setRightBorderColor(color.getIndex());
                        }
                        break;
                    case Y:
                        // 设置上下边框
                        cellStyle.setBorderTop(style);
                        cellStyle.setBorderBottom(style);
                        if (color != null) {
                            cellStyle.setTopBorderColor(color.getIndex());
                            cellStyle.setBottomBorderColor(color.getIndex());
                        }
                        break;
                    case TOP:
                        cellStyle.setBorderTop(style);
                        if (color != null)
                            cellStyle.setTopBorderColor(color.getIndex());
                        break;
                    case BOTTOM:
                        cellStyle.setBorderBottom(style);
                        if (color != null)
                            cellStyle.setBottomBorderColor(color.getIndex());
                        break;
                    case LEFT:
                        cellStyle.setBorderLeft(style);
                        if (color != null)
                            cellStyle.setLeftBorderColor(color.getIndex());
                        break;
                    case RIGHT:
                        cellStyle.setBorderRight(style);
                        if (color != null)
                            cellStyle.setRightBorderColor(color.getIndex());
                        break;
                }
            }
        }

        // 设置对齐方式
        VerticalAlignment verticalAlignment = excelColumn.getVerticalAlignment();
        if (verticalAlignment != null)
            cellStyle.setVerticalAlignment(verticalAlignment);

        HorizontalAlignment horizontalAlignment = excelColumn.getHorizontalAlignment();
        if (horizontalAlignment != null)
            cellStyle.setAlignment(horizontalAlignment);

        cell.setCellStyle(cellStyle);
    }

    /**
     * 设置列宽
     * 
     * @param sheet 工作表
     * @param index 列索引
     * @param width 列宽
     */
    private static void setCellWidth(Sheet sheet, int index, Integer width) {
        if (width == null || width <= 0)
            return;
        sheet.setColumnWidth(index, width);
    }

    /**
     * 设置行高
     * 
     * @param row    行
     * @param height 行高
     */
    private static void setCellHeight(Row row, Short height) {
        if (height == null || height <= 0)
            return;
        row.setHeight(height);
    }

    /**
     * 保存Excel文件
     * <p>
     * 将工作簿保存到指定路径，自动生成唯一文件名
     * </p>
     * 
     * @param filePath 文件路径（相对路径）
     * @param workbook Excel工作簿
     * @return 生成的文件路径
     * @throws RuntimeException 当路径无效或写入失败时抛出
     */
    public static String toFile(String filePath, Workbook workbook) {
        if (filePath.startsWith("/"))
            throw new RuntimeException("暂不支持绝对路径");

        String[] filePaths = Arrays.stream(filePath.split("/"))
                .filter(path -> path != null && !path.trim().isEmpty())
                .toArray(String[]::new);

        if (ArrayUtils.isEmpty(filePaths))
            filePaths = new String[] { "excel" };

        // 生成唯一文件名
        String file = String.format("%s_%d.xlsx", UUID.randomUUID(), System.currentTimeMillis());
        String fullPath = String.join(File.separator, filePaths) + File.separator + file;

        try (OutputStream out = getOutputStream(fullPath)) {
            workbook.write(out);
            out.flush();
            return String.join("/", filePaths) + "/" + file;
        } catch (IOException e) {
            throw new RuntimeException("文件写入失败: " + e.getMessage(), e);
        }
    }

    /**
     * 获取文件输出流
     * <p>
     * 如果目录不存在会自动创建
     * </p>
     * 
     * @param path 文件路径
     * @return 文件输出流
     * @throws IOException IO异常
     */
    private static FileOutputStream getOutputStream(String path) throws IOException {
        File file = new File(path);
        // 创建父目录
        if (!file.getParentFile().exists()) {
            file.getParentFile().mkdirs();
        }
        // 创建文件
        if (!file.exists()) {
            file.createNewFile();
        }
        return new FileOutputStream(file, false);
    }
}
