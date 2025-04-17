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
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Array;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

public class ExcelReaderUtils {

    public static <T> T doIt(Sheet sheet, Class<T> clazz) {
        if (sheet == null) throw new RuntimeException("表格数据不能为空");
        T obj;
        try {
            obj = clazz.getDeclaredConstructor().newInstance();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        List<FieldCell> fieldCellList = getFieldCells(clazz, null, null, null);

        for (FieldCell fieldCell : fieldCellList) {
            if (fieldCell.getRowIndex() == -1 || fieldCell.getStartCellIndex() == -1) continue;
            Row row = sheet.getRow(fieldCell.getRowIndex());
            if (row == null) continue;
            Cell cell = row.getCell(fieldCell.getStartCellIndex());
            if (cell == null) continue;
            ObjectUtils.setField(obj, fieldCell.getField(), getValue(fieldCell.getField().getType(), cell, fieldCell.getFormatter()));
        }

        return obj;
    }

    public static <T> List<T> doList(Sheet sheet, Class<T> clazz, IExcelReaderListener doExcel) {
        if (sheet == null) throw new RuntimeException("表格数据不能为空");
        int startHeaderNumber;
        int endHeaderNumber;
        if (doExcel != null) {
            startHeaderNumber = doExcel.startHeaderNumber(sheet) - 1;
            endHeaderNumber = doExcel.headerNumber(sheet);
        } else {
            startHeaderNumber = 0;
            endHeaderNumber = 0;
        }

        if (startHeaderNumber < 0) throw new RuntimeException("表头开始行数不正确");
        if (endHeaderNumber < 0) throw new RuntimeException("表头行数不能小于0");

        List<FieldCell> fieldCellList = getFieldCells(clazz, sheet, startHeaderNumber, endHeaderNumber);

        return IntStream.range(endHeaderNumber, sheet.getLastRowNum() + 1).mapToObj(rowIndex -> {
            Row row = sheet.getRow(rowIndex);
            if (doExcel != null && doExcel.isFooter(row)) return null;
            T obj;
            try {
                obj = clazz.getDeclaredConstructor().newInstance();
            } catch (Exception e) {
                throw new RuntimeException(e);
            }

            for (FieldCell fieldCell : fieldCellList) {
                if (fieldCell.getRowIndex() == null && fieldCell.getStartCellIndex() == null)
                    continue;
                if (Objects.equals(-1, fieldCell.getRowIndex()) || Objects.equals(-1, fieldCell.getStartCellIndex()))
                    continue;

                Class<?> typeClazz = fieldCell.getField().getType();
                if (Map.class.isAssignableFrom(typeClazz)) {
                    throw new RuntimeException("暂不支持Map类型");
                } else if (typeClazz.isEnum()) {
                    if (fieldCell.getFormatter() == null) throw new RuntimeException("请自定义转换器");
                    if (!Objects.equals(fieldCell.getStartCellIndex(), fieldCell.getEndCellIndex()))
                        throw new RuntimeException("枚举类型不支持多列");
                    Cell cell = row.getCell(fieldCell.getStartCellIndex());
                    if (cell == null) continue;
                    fieldCell.getFormatter().format(cell);
                } else if (typeClazz.isArray()) {
                    // 获取数组元素类型
                    Class<?> componentType = typeClazz.getComponentType();
                    // 处理数组
                    List<Object> list = IntStream.range(fieldCell.getStartCellIndex(), fieldCell.getEndCellIndex() + 1).mapToObj(cellIndex -> {
                        Cell cell = row.getCell(cellIndex);
                        return getValue(componentType, cell, fieldCell.getFormatter());
                    }).collect(Collectors.toList());

                    // 创建并填充数组
                    Object array = Array.newInstance(componentType, list.size());
                    for (int i = 0; i < list.size(); i++) {
                        Array.set(array, i, list.get(i));
                    }
                    ObjectUtils.setField(obj, fieldCell.getField(), array);
                } else if (Collection.class.isAssignableFrom(typeClazz)) {
                    Type rowType = ((ParameterizedType) fieldCell.getField().getGenericType()).getRawType();
                    Type[] types = ((ParameterizedType) fieldCell.getField().getGenericType()).getActualTypeArguments();
                    if (types.length != 1) throw new RuntimeException("未知类型");
                    Class<?> typeClass = (Class<?>) types[0];

                    Stream<Object> stream = IntStream.range(fieldCell.getStartCellIndex(), fieldCell.getEndCellIndex() + 1).mapToObj(cellIndex -> {
                        Cell cell = row.getCell(cellIndex);
                        return getValue(typeClass, cell, fieldCell.getFormatter());
                    });

                    if (rowType instanceof List<?>) {
                        ObjectUtils.setField(obj, fieldCell.getField(), stream.collect(Collectors.toList()));
                    } else if (rowType instanceof Set<?>) {
                        ObjectUtils.setField(obj, fieldCell.getField(), stream.collect(Collectors.toSet()));
                    } else {
                        throw new RuntimeException("暂不支持该数据类型");
                    }
                } else {
                    if (!Objects.equals(fieldCell.getStartCellIndex(), fieldCell.getEndCellIndex()))
                        throw new RuntimeException("该数据类型不支持多列");
                    Cell cell = row.getCell(fieldCell.getStartCellIndex());
                    if (cell == null) continue;
                    ObjectUtils.setField(obj, fieldCell.getField(), getValue(typeClazz, cell, fieldCell.getFormatter()));
                }
            }
            return obj;
        }).filter(ObjectUtils::isNotEmpty).collect(Collectors.toList());

    }

    public static List<HeaderCell> getHeaders(Sheet sheet, boolean isSingle, IExcelReaderListener doExcel) {
        if (sheet == null) throw new RuntimeException("表格数据不能为空");
        int startHeaderNumber;
        int endHeaderNumber;
        if (doExcel != null) {
            startHeaderNumber = doExcel.startHeaderNumber(sheet) - 1;
            endHeaderNumber = doExcel.headerNumber(sheet);
        } else {
            startHeaderNumber = 0;
            endHeaderNumber = 0;
        }
        if (startHeaderNumber < 0) throw new RuntimeException("表头开始行数不正确");
        if (endHeaderNumber < 0) throw new RuntimeException("表头行数不能小于0");

        List<CellRangeAddress> mergedRegionList = sheet.getMergedRegions().stream().filter(it -> it.getFirstRow() <= endHeaderNumber).collect(Collectors.toList());
        List<HeaderCell> headerCellList = getHeaderCellList(sheet, startHeaderNumber, endHeaderNumber, mergedRegionList);
        if (!isSingle) return headerCellList;
        List<HeaderCell> finalHeaderCellList;
        if (endHeaderNumber - startHeaderNumber > 1) {

            int startCellIndex = headerCellList.stream().mapToInt(HeaderCell::getStartCellIndex).min().orElse(0);
            int endCellIndex = headerCellList.stream().mapToInt(HeaderCell::getEndCellIndex).max().orElse(0);

            finalHeaderCellList = IntStream.range(startCellIndex, endCellIndex + 1).mapToObj(index -> {
                List<HeaderCell> list = headerCellList.stream().filter(it -> it.getStartCellIndex() <= index && it.getEndCellIndex() >= index).sorted(Comparator.comparing(HeaderCell::getRowIndex)).collect(Collectors.toList());

                if (CollectionUtils.isEmpty(list)) return null;
                if (list.size() == 1) return list.get(0);
                else {
                    String cellValue = list.stream().map(it -> it.getCellValue() == null ? "" : it.getCellValue().toString()).collect(Collectors.joining("-"));
                    HeaderCell last = list.get(list.size() - 1);
                    last.setCellValue(cellValue);
                    return last;
                }
            }).collect(Collectors.toList());
        } else finalHeaderCellList = headerCellList;
        return finalHeaderCellList;
    }

    public static List<Map<String, Object>> doMap(Sheet sheet, List<HeaderCell> headerCellList, IExcelReaderListener doExcel) {
        if (sheet == null) throw new RuntimeException("表格数据不能为空");
        if (CollectionUtils.isEmpty(headerCellList)) throw new RuntimeException("表头数据不能为空");
        headerCellList.sort(Comparator.comparing(HeaderCell::getStartCellIndex).thenComparing(HeaderCell::getEndCellIndex));
        int headerNumber;
        if (doExcel != null) {
            headerNumber = doExcel.headerNumber(sheet);
        } else {
            headerNumber = 0;
        }
        if (headerNumber < 0) throw new RuntimeException("表头行数不能小于0");

        return IntStream.range(headerNumber, sheet.getLastRowNum() + 1).mapToObj(rowIndex -> {
            Row row = sheet.getRow(rowIndex);
            if (doExcel != null && doExcel.isFooter(row)) return null;
            Map<String, Object> map = new LinkedHashMap<>();
            for (HeaderCell headerCell : headerCellList) {
                Object obj;
                if (headerCell.getStartCellIndex() == headerCell.getEndCellIndex()) {
                    obj = getCellValue(row.getCell(headerCell.getStartCellIndex()));
                } else {
                    obj = IntStream.range(headerCell.getStartCellIndex(), headerCell.getEndCellIndex() + 1).mapToObj(cellIndex -> {
                        Cell cell = row.getCell(cellIndex);
                        return getCellValue(cell);
                    }).collect(Collectors.toList());
                }
                map.put(headerCell.getCellValue().toString(), obj);
            }
            return map;
        }).collect(Collectors.toList());
    }

    public static List<Map<String, Object>> doMap(Sheet sheet, IExcelReaderListener doExcel) {
        if (sheet == null) throw new RuntimeException("表格数据不能为空");
        List<HeaderCell> headerCellList = getHeaders(sheet, true, doExcel);
        if (CollectionUtils.isEmpty(headerCellList)) throw new RuntimeException("表头数据不能为空");
        headerCellList.sort(Comparator.comparing(HeaderCell::getStartCellIndex).thenComparing(HeaderCell::getEndCellIndex));
        int headerNumber;
        if (doExcel != null) {
            headerNumber = doExcel.headerNumber(sheet);
        } else {
            headerNumber = 0;
        }
        if (headerNumber < 0) throw new RuntimeException("表头行数不能小于0");

        return IntStream.range(headerNumber, sheet.getLastRowNum() + 1).mapToObj(rowIndex -> {
            Row row = sheet.getRow(rowIndex);
            if (doExcel != null && doExcel.isFooter(row)) return null;
            Map<String, Object> map = new LinkedHashMap<>();
            for (HeaderCell headerCell : headerCellList) {
                Object obj;
                if (headerCell.getStartCellIndex() == headerCell.getEndCellIndex()) {
                    obj = getCellValue(row.getCell(headerCell.getStartCellIndex()));
                } else {
                    obj = IntStream.range(headerCell.getStartCellIndex(), headerCell.getEndCellIndex() + 1).mapToObj(cellIndex -> {
                        Cell cell = row.getCell(cellIndex);
                        return getCellValue(cell);
                    }).collect(Collectors.toList());
                }
                map.put(headerCell.getCellValue().toString(), obj);
            }
            return map;
        }).collect(Collectors.toList());
    }

    private static List<FieldCell> getFieldCells(Class<?> clazz, Sheet sheet, Integer startHeaderNumber, Integer endHeaderNumber) {
        Field[] fields = clazz.getDeclaredFields();
        return Arrays.stream(fields).map(field -> {
            FieldCell fieldCell = new FieldCell();
            fieldCell.setField(field);
            if (!field.isAnnotationPresent(ExcelReader.class)) return fieldCell;
            ExcelReader annotation = field.getAnnotation(ExcelReader.class);
            if (annotation == null) return fieldCell;
            if (annotation.formatter() != DefaultExcelReaderFormatter.class) {
                try {
                    fieldCell.setFormatter(annotation.formatter().getDeclaredConstructor().newInstance());
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            }
            String cellValue = annotation.value();
            int cellIndex = annotation.index();
            String[] cellNames = annotation.name();
            if (StringUtils.isNotEmpty(cellValue)) {
                // 判断坐标
                Matcher matcher = Pattern.compile("(\\D+)(\\d+)").matcher(cellValue);
                if (matcher.find()) {
                    String[] parts = {matcher.group(1), matcher.group(2)};
                    fieldCell.setRowIndex(Integer.parseInt(parts[1]) - 1);
                    fieldCell.setStartCellIndex(ObjectUtils.convertToNumber(parts[0]) - 1);
                    fieldCell.setEndCellIndex(ObjectUtils.convertToNumber(parts[0]) - 1);
                }
            } else if (cellIndex != -1) {
                // 判断游标
                fieldCell.setStartCellIndex(cellIndex);
                fieldCell.setEndCellIndex(cellIndex);
            } else if (ArrayUtils.isNotEmpty(cellNames)) {
                List<HeaderCell> headerCellList;
                List<CellRangeAddress> mergedRegionList;
                if (sheet != null) {
                    mergedRegionList = sheet.getMergedRegions().stream().filter(it -> it.getFirstRow() <= endHeaderNumber).collect(Collectors.toList());
                    headerCellList = getHeaderCellList(sheet, startHeaderNumber, endHeaderNumber, mergedRegionList);
                } else headerCellList = null;
                // 判断名称
                if (headerCellList == null) throw new RuntimeException("无法解析表头数据");
                HeaderCell tmpHeaderCell = null;
                for (int i = 0; i < cellNames.length; i++) {
                    String name = cellNames[i];
                    HeaderCell finalTmpHeaderCell = tmpHeaderCell;
                    List<HeaderCell> list = headerCellList.stream().filter(it -> {
                        // 多表头的情况，且父表头已获取到数据时
                        if (finalTmpHeaderCell != null) {
                            // 判断子表头肯定在父表头的下一行
                            if (finalTmpHeaderCell.getRowIndex() >= it.getRowIndex()) return false;
                            // 判断子表头列下标需在父表头的内部
                            return finalTmpHeaderCell.getStartCellIndex() <= it.getStartCellIndex() && finalTmpHeaderCell.getEndCellIndex() >= it.getEndCellIndex() && Objects.equals(name, it.getCellValue());
                        }
                        return Objects.equals(name, it.getCellValue());
                    }).collect(Collectors.toList());
                    /// 没有对应数据时
                    if (CollectionUtils.isEmpty(list)) {
                        // 如果最后的数据匹配不到就break
                        if (i == cellNames.length - 1) {
                            tmpHeaderCell = null;
                            break;
                        } else {
                            // 继续尝试获取子表头的数据
                            continue;
                        }
                    }

                    // 如果有数据
                    if (list.size() == 1)
                        // 单条数据
                        tmpHeaderCell = list.get(0);
                    else {
                        // 多条数据
                        if (i != cellNames.length - 1) {
                            // 父表头有重复数据时
                            throw new RuntimeException("表头过于复杂，推荐使用 @ExcelValue(index = ?) 方式处理数据");
                        } else {
                            // 过滤掉有父表头的数据
                            list = list.stream().filter(it -> headerCellList.stream().filter(headerCell -> headerCell.getRowIndex() != it.getRowIndex() && headerCell.getStartCellIndex() <= it.getStartCellIndex() && headerCell.getEndCellIndex() >= it.getEndCellIndex()).count() == 1).collect(Collectors.toList());
                            if (CollectionUtils.isEmpty(list))
                                // emmm 表头数据重复？
                                throw new RuntimeException("表头过于复杂，推荐使用 @ExcelValue(index = ?) 方式处理数据");
                            else if (list.size() == 1)
                                // 单条数据
                                tmpHeaderCell = list.get(0);
                            else
                                // 子表头有重复数据
                                throw new RuntimeException("表头过于复杂，推荐使用 @ExcelValue(index = ?) 方式处理数据");

                        }
                    }
                }
                if (tmpHeaderCell != null) {
                    fieldCell.setStartCellIndex(tmpHeaderCell.getStartCellIndex());
                    fieldCell.setEndCellIndex(tmpHeaderCell.getEndCellIndex());
                }
            }

            return fieldCell;

        }).collect(Collectors.toList());
    }

    private static List<HeaderCell> getHeaderCellList(Sheet sheet, Integer startHeaderNumber, Integer endHeaderNumber, List<CellRangeAddress> mergedRegionList) {
        if (startHeaderNumber >= endHeaderNumber) return Collections.emptyList();

        return IntStream.range(startHeaderNumber, endHeaderNumber).mapToObj(rowIndex -> {
            Row row = sheet.getRow(rowIndex);
            return IntStream.range(0, row.getLastCellNum()).mapToObj(cellIndex -> {
                Cell cell = row.getCell(cellIndex);
                Object cellValue = getCellValue(cell);
                CellRangeAddress cellAddresses = mergedRegionList.stream().filter(range -> range.getFirstRow() <= rowIndex && range.getLastRow() >= rowIndex && range.getFirstColumn() <= cellIndex && range.getLastColumn() >= cellIndex).findFirst().orElse(null);
                if (cellAddresses != null) {
                    if (Objects.equals(cellAddresses.getFirstRow(), rowIndex) && Objects.equals(cellAddresses.getFirstColumn(), cellIndex))
                        return new HeaderCell(cellValue, rowIndex, cellIndex, cellAddresses.getLastColumn());
                    else return null;
                } else return new HeaderCell(cellValue, rowIndex, cellIndex, cellIndex);
            }).filter(it -> it != null && it.getCellValue() != null).collect(Collectors.toList());
        }).flatMap(Collection::stream).collect(Collectors.toList());
    }

    public static Object getValue(Class<?> clazz, Cell cell, IExcelReaderFormatter<?> formatter) {
        if (cell == null) return null;
        if (formatter != null) return formatter.format(cell);
        if (clazz.equals(String.class)) {
            Object object = getCellValue(cell);
            return object == null ? null : object.toString();
        }
        if (Number.class.isAssignableFrom(clazz)) {
            try {
                double value = cell.getNumericCellValue();
                if ((clazz.equals(Integer.class) || clazz.equals(Long.class)) && (String.valueOf(value).matches(".*\\.\\d*[1-9]+\\d*$")))
                    throw new RuntimeException("当前数据是浮点类型");
                BigDecimal cellValue = BigDecimal.valueOf(value);
                if (clazz.equals(Integer.class)) return cellValue.intValue();
                if (clazz.equals(Double.class)) return cellValue.doubleValue();
                if (clazz.equals(Float.class)) return cellValue.floatValue();
                if (clazz.equals(Long.class)) return cellValue.longValue();
            } catch (IllegalStateException e) {
                if (!Objects.equals(CellType.STRING, cell.getCellType())) throw e;
                String cellValue = cell.getStringCellValue();
                if (StringUtils.isEmpty(cellValue)) return null;
                if (clazz.equals(Integer.class)) return Integer.parseInt(cellValue);
                if (clazz.equals(Double.class)) return Double.parseDouble(cellValue);
                if (clazz.equals(Float.class)) return Float.parseFloat(cellValue);
                if (clazz.equals(Long.class)) return Long.parseLong(cellValue);
                return null;
            }
        }
        if (clazz.equals(Date.class)) return new SimpleDateExcelReaderFormatter().format(cell);
        if (clazz.equals(Boolean.class)) return cell.getBooleanCellValue();
        throw new RuntimeException("暂不支持该格式");
    }

    public static Object getCellValue(Cell cell) {
        if (cell == null) return null;

        Object obj;
        switch (cell.getCellType()) {
            case STRING:
                obj = cell.getStringCellValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) obj = cell.getDateCellValue();
                else obj = cell.getNumericCellValue();
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
