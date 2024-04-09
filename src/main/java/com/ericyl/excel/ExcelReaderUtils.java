package com.ericyl.excel;


import com.ericyl.excel.reader.IExcelReaderListener;
import com.ericyl.excel.reader.annotation.ExcelReader;
import com.ericyl.excel.reader.formatter.DefaultExcelReaderFormatter;
import com.ericyl.excel.reader.formatter.IExcelReaderFormatter;
import com.ericyl.excel.reader.formatter.SimpleDateExcelReaderFormatter;
import com.ericyl.excel.reader.model.FieldCell;
import com.ericyl.excel.reader.model.HeaderCell;
import com.ericyl.excel.util.ObjectUtils;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class ExcelReaderUtils {

    public static <T> T doIt(Sheet sheet, Class<T> clazz) {
        if (sheet == null)
            throw new RuntimeException("表格数据不能为空");
        T obj;
        try {
            obj = clazz.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }

        List<FieldCell> fieldCellList = getFieldCells(clazz, null, null);

        IntStream.range(0, sheet.getLastRowNum() + 1).forEach(rowIndex -> {
            Row row = sheet.getRow(rowIndex);

            List<FieldCell> rowFieldCellList = fieldCellList.stream().filter(fieldCell -> Objects.equals(fieldCell.getRowIndex(), rowIndex + 1)).collect(Collectors.toList());
            if (CollectionUtils.isNotEmpty(rowFieldCellList)) {

                IntStream.range(0, row.getLastCellNum()).forEach(cellIndex -> {
                    rowFieldCellList.stream().filter(it -> Objects.equals(it.getCellIndex(), cellIndex + 1)).findFirst().ifPresent(fieldCell -> ObjectUtils.setField(obj, fieldCell.getField(), getValue(fieldCell.getField().getType(), row.getCell(cellIndex), fieldCell.getFormatter())));
                });
            }
        });

        return obj;
    }

    public static <T> List<T> doList(Sheet sheet, Class<T> clazz, IExcelReaderListener doExcel) {
        if (sheet == null)
            throw new RuntimeException("表格数据不能为空");
        int headerNumber = 0;
        if (doExcel != null) {
            headerNumber = doExcel.headerNumber(sheet);
        }
        if (headerNumber < 0)
            throw new RuntimeException("表头行数不能小于0");

        List<FieldCell> fieldCellList = getFieldCells(clazz, sheet, headerNumber);

        return IntStream.range(headerNumber, sheet.getLastRowNum() + 1).mapToObj(rowIndex -> {
            Row row = sheet.getRow(rowIndex);
            if (doExcel != null && doExcel.isFooter(row))
                return null;
            T obj;
            try {
                obj = clazz.newInstance();
            } catch (InstantiationException | IllegalAccessException e) {
                throw new RuntimeException(e);
            }

            IntStream.range(0, row.getLastCellNum()).forEach(cellIndex -> fieldCellList.stream().filter(it -> Objects.equals(it.getCellIndex(), cellIndex)).findFirst().ifPresent(fieldCell -> ObjectUtils.setField(obj, fieldCell.getField(), getValue(fieldCell.getField().getType(), row.getCell(cellIndex), fieldCell.getFormatter()))));

            return obj;
        }).filter(ObjectUtils::isNotEmpty).collect(Collectors.toList());

    }

    private static List<FieldCell> getFieldCells(Class<?> clazz, Sheet sheet, Integer headerNumber) {
        List<HeaderCell> headerCellList;
        List<CellRangeAddress> mergedRegionList;
        if (sheet != null) {
            mergedRegionList = sheet.getMergedRegions().stream().filter(it -> it.getFirstRow() <= headerNumber).collect(Collectors.toList());
            headerCellList = getHeaderCellList(sheet, headerNumber, mergedRegionList);
        } else {
            headerCellList = null;
        }
        Field[] fields = clazz.getDeclaredFields();
        return Arrays.stream(fields).map(field -> {
            FieldCell fieldCell = new FieldCell();
            fieldCell.setField(field);
            if (!field.isAnnotationPresent(ExcelReader.class))
                return fieldCell;
            ExcelReader annotation = field.getAnnotation(ExcelReader.class);
            if (annotation.formatter() != DefaultExcelReaderFormatter.class)
                try {
                    fieldCell.setFormatter(annotation.formatter().newInstance());
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            String cellValue = annotation.value();
            int cellIndex = annotation.index();
            String[] cellNames = annotation.name();
            if (StringUtils.isNotEmpty(cellValue)) {
                //判断坐标
                Matcher matcher = Pattern.compile("(\\D+)(\\d+)").matcher(cellValue);
                if (matcher.find()) {
                    String[] parts = {matcher.group(1), matcher.group(2)};
                    fieldCell.setRowIndex(Integer.parseInt(parts[1]));
                    fieldCell.setCellIndex(ObjectUtils.convertToNumber(parts[0]));
                }
            } else if (cellIndex != -1) {
                //判断游标
                fieldCell.setCellIndex(cellIndex);
            } else if (ArrayUtils.isNotEmpty(cellNames)) {
                //判断名称
                if (headerCellList == null)
                    throw new RuntimeException("无法解析表头数据");
                HeaderCell tmpHeaderCell = null;
                for (int i = 0; i < cellNames.length; i++) {
                    String name = cellNames[i];
                    HeaderCell finalTmpHeaderCell = tmpHeaderCell;
                    List<HeaderCell> list = headerCellList.stream().filter(it -> {
                        //多表头的情况，且父表头已获取到数据时
                        if (finalTmpHeaderCell != null) {
                            //判断子表头肯定在父表头的下一行
                            if (finalTmpHeaderCell.getRowIndex() >= it.getRowIndex())
                                return false;
                            //判断子表头列下标需在父表头的内部
                            return finalTmpHeaderCell.getStartCellIndex() <= it.getStartCellIndex() && finalTmpHeaderCell.getEndCellIndex() >= it.getEndCellIndex() && Objects.equals(name, it.getCellValue());
                        }
                        return Objects.equals(name, it.getCellValue());
                    }).collect(Collectors.toList());
                    ///没有对应数据时
                    if (CollectionUtils.isEmpty(list)) {
                        //如果最后的数据匹配不到就break
                        if (i == cellNames.length - 1) {
                            tmpHeaderCell = null;
                            break;
                        } else {
                            //继续尝试获取子表头的数据
                            continue;
                        }
                    }

                    //如果有数据
                    if (list.size() == 1)
                        //单条数据
                        tmpHeaderCell = list.get(0);
                    else {
                        //多条数据
                        if (i != cellNames.length - 1) {
                            //父表头有重复数据时
                            throw new RuntimeException("表头过于复杂，推荐使用 @ExcelValue(index = ?) 方式处理数据");
                        } else {
                            //过滤掉有父表头的数据
                            list = list.stream()
                                    .filter(it -> headerCellList.stream().filter(headerCell -> headerCell.getRowIndex() != it.getRowIndex() && headerCell.getStartCellIndex() <= it.getStartCellIndex() && headerCell.getEndCellIndex() >= it.getEndCellIndex()).count() == 1)
                                    .collect(Collectors.toList());
                            if (CollectionUtils.isEmpty(list))
                                //emmm 表头数据重复？
                                throw new RuntimeException("表头过于复杂，推荐使用 @ExcelValue(index = ?) 方式处理数据");
                            else if (list.size() == 1)
                                //单条数据
                                tmpHeaderCell = list.get(0);
                            else
                                //子表头有重复数据
                                throw new RuntimeException("表头过于复杂，推荐使用 @ExcelValue(index = ?) 方式处理数据");

                        }
                    }
                }
                if (tmpHeaderCell != null)
                    fieldCell.setCellIndex(tmpHeaderCell.getStartCellIndex());
            }

            return fieldCell;

        }).collect(Collectors.toList());
    }

    private static List<HeaderCell> getHeaderCellList(Sheet sheet, Integer headerNumber, List<CellRangeAddress> mergedRegionList) {
        return IntStream.range(0, headerNumber).mapToObj(rowIndex -> {
            Row row = sheet.getRow(rowIndex);
            return IntStream.range(0, row.getLastCellNum()).mapToObj(cellIndex -> {
                Cell cell = row.getCell(cellIndex);
                Object cellValue = getCellValue(cell);
                CellRangeAddress cellAddresses = mergedRegionList.stream().filter(range -> range.getFirstRow() <= rowIndex && range.getLastRow() >= rowIndex && range.getFirstColumn() <= cellIndex && range.getLastColumn() >= cellIndex).findFirst().orElse(null);
                if (cellAddresses != null) {
                    if (Objects.equals(cellAddresses.getFirstRow(), rowIndex) && Objects.equals(cellAddresses.getFirstColumn(), cellIndex))
                        return new HeaderCell(cellValue, rowIndex, cellIndex, cellAddresses.getLastColumn());
                    else return null;
                } else
                    return new HeaderCell(cellValue, rowIndex, cellIndex, cellIndex);
            }).filter(it -> it != null && it.getCellValue() != null).collect(Collectors.toList());
        }).flatMap(Collection::stream).collect(Collectors.toList());
    }

    private static Object getValue(Class<?> clazz, Cell cell, IExcelReaderFormatter<?> formatter) {
        if (cell == null)
            return null;
        if (formatter != null)
            return formatter.format(cell);
        if (clazz.equals(String.class))
            return cell.getStringCellValue();
        if (Number.class.isAssignableFrom(clazz)) {
            try {
                return cell.getNumericCellValue();
            } catch (IllegalStateException e) {
                if (!Objects.equals(CellType.STRING, cell.getCellType()))
                    throw e;
                String cellValue = cell.getStringCellValue();
                if (StringUtils.isEmpty(cellValue))
                    return null;
                if (clazz.equals(Integer.class))
                    return Integer.parseInt(cellValue);
                if (clazz.equals(Double.class))
                    return Double.parseDouble(cellValue);
                if (clazz.equals(Float.class))
                    return Float.parseFloat(cellValue);
                if (clazz.equals(Long.class))
                    return Long.parseLong(cellValue);
                return null;
            }
        }
        if (clazz.equals(Date.class))
            return new SimpleDateExcelReaderFormatter().format(cell);
        if (clazz.equals(Boolean.class))
            return cell.getBooleanCellValue();
        throw new RuntimeException("暂不支持该格式");
    }

    public static Object getCellValue(Cell cell) {
        if (cell == null)
            return null;
        Object obj;
        switch (cell.getCellType()) {
            case STRING:
                obj = cell.getStringCellValue();
                break;
            case NUMERIC:
                obj = cell.getNumericCellValue();
                break;
            case BOOLEAN:
                obj = cell.getBooleanCellValue();
                break;
            case _NONE:
            case FORMULA:
            case BLANK:
            case ERROR:
            default:
                obj = null;
                break;
        }

        return obj;
    }


}

