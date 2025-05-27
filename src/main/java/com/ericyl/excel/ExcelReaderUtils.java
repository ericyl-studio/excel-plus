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

/**
 * Excel 读取工具类
 * <p>
 * 提供多种方式读取 Excel 数据：
 * 1. 单对象读取：通过坐标定位读取特定单元格数据到对象
 * 2. 列表读取：通过表头或索引批量读取数据到列表
 * 3. Map读取：将表格数据读取为 Map 格式
 * </p>
 *
 * @author ericyl
 * @since 1.0
 */
public class ExcelReaderUtils {

    /**
     * 坐标正则表达式，用于匹配 Excel 坐标格式（如 A1, B2）
     */
    private static final Pattern COORDINATE_PATTERN = Pattern.compile("(\\D+)(\\d+)");

    /**
     * 多行表头连接符
     */
    private static final String HEADER_SEPARATOR = "-";

    /**
     * 读取单个对象数据
     * <p>
     * 通过 @ExcelReader 注解中的坐标值（如 "A1"）定位并读取单元格数据，
     * 自动映射到指定类型的对象字段中
     * </p>
     *
     * @param sheet Excel工作表
     * @param clazz 目标对象类型
     * @param <T>   泛型类型
     * @return 填充数据后的对象实例
     * @throws RuntimeException 当表格数据为空或对象创建失败时抛出
     */
    public static <T> T doIt(Sheet sheet, Class<T> clazz) {
        if (sheet == null)
            throw new RuntimeException("表格数据不能为空");
        T obj;
        try {
            obj = clazz.getDeclaredConstructor().newInstance();
        } catch (Exception e) {
            throw new RuntimeException("创建对象实例失败: " + e.getMessage(), e);
        }

        List<FieldCell> fieldCellList = getFieldCells(clazz, null, null, null);

        for (FieldCell fieldCell : fieldCellList) {
            if (fieldCell.getRowIndex() == -1 || fieldCell.getStartCellIndex() == -1)
                continue;
            ObjectUtils.setField(obj, fieldCell.getField(),
                    getValue(fieldCell.getField().getType(), sheet, fieldCell.getRowIndex(),
                            fieldCell.getStartCellIndex(), fieldCell.getFormatter()));
        }

        return obj;
    }

    /**
     * 读取列表数据
     * <p>
     * 支持通过以下方式定位数据：
     * 1. 索引方式：@ExcelReader(index = 0)
     * 2. 表头方式：@ExcelReader(name = {"表头1", "表头2"})
     * 支持多表头和复杂数据类型（数组、集合等）
     * </p>
     *
     * @param sheet   Excel工作表
     * @param clazz   列表元素类型
     * @param doExcel Excel读取监听器，用于自定义表头和表尾判断逻辑
     * @param <T>     泛型类型
     * @return 数据列表
     * @throws RuntimeException 当表格数据为空或数据处理失败时抛出
     */
    public static <T> List<T> doList(Sheet sheet, Class<T> clazz, IExcelReaderListener doExcel) {
        if (sheet == null)
            throw new RuntimeException("表格数据不能为空");

        // 获取表头行范围
        int startHeaderNumber;
        int endHeaderNumber;
        if (doExcel != null) {
            startHeaderNumber = doExcel.startHeaderNumber(sheet) - 1;
            endHeaderNumber = doExcel.endHeaderNumber(sheet);
        } else {
            startHeaderNumber = 0;
            endHeaderNumber = 0;
        }

        if (startHeaderNumber < 0)
            throw new RuntimeException("表头开始行数不正确");
        if (endHeaderNumber < 0)
            throw new RuntimeException("表头行数不能小于0");

        // 解析字段与单元格的映射关系
        List<FieldCell> fieldCellList = getFieldCells(clazz, sheet, startHeaderNumber, endHeaderNumber);

        // 逐行读取数据
        return IntStream.range(endHeaderNumber, sheet.getLastRowNum() + 1).mapToObj(rowIndex -> {
            Row row = sheet.getRow(rowIndex);
            // 跳过空行
            if (row == null)
                return null;
            // 跳过表尾
            if (doExcel != null && doExcel.isFooter(row))
                return null;

            T obj;
            try {
                obj = clazz.getDeclaredConstructor().newInstance();
            } catch (Exception e) {
                throw new RuntimeException("创建对象实例失败: " + e.getMessage(), e);
            }

            // 处理每个字段
            for (FieldCell fieldCell : fieldCellList) {
                if (fieldCell.getRowIndex() == null && fieldCell.getStartCellIndex() == null)
                    continue;
                if (Objects.equals(-1, fieldCell.getRowIndex()) || Objects.equals(-1, fieldCell.getStartCellIndex()))
                    continue;

                Class<?> typeClazz = fieldCell.getField().getType();

                // 根据字段类型进行不同的处理
                if (Map.class.isAssignableFrom(typeClazz)) {
                    throw new RuntimeException("暂不支持Map类型");
                } else if (typeClazz.isEnum()) {
                    // 枚举类型处理
                    if (fieldCell.getFormatter() == null)
                        throw new RuntimeException("枚举类型请自定义转换器");
                    if (!Objects.equals(fieldCell.getStartCellIndex(), fieldCell.getEndCellIndex()))
                        throw new RuntimeException("枚举类型不支持多列");
                    Cell cell = row.getCell(fieldCell.getStartCellIndex());
                    if (cell == null)
                        continue;
                    Object enumValue = fieldCell.getFormatter().format(cell);
                    ObjectUtils.setField(obj, fieldCell.getField(), enumValue);
                } else if (typeClazz.isArray()) {
                    // 数组类型处理
                    Class<?> componentType = typeClazz.getComponentType();
                    List<Object> list = IntStream.range(fieldCell.getStartCellIndex(), fieldCell.getEndCellIndex() + 1)
                            .mapToObj(cellIndex -> getValue(componentType, sheet, rowIndex, cellIndex, fieldCell.getFormatter())).collect(Collectors.toList());

                    // 创建并填充数组
                    Object array = Array.newInstance(componentType, list.size());
                    for (int i = 0; i < list.size(); i++) {
                        Array.set(array, i, list.get(i));
                    }
                    ObjectUtils.setField(obj, fieldCell.getField(), array);
                } else if (Collection.class.isAssignableFrom(typeClazz)) {
                    // 集合类型处理
                    Type genericType = fieldCell.getField().getGenericType();
                    if (!(genericType instanceof ParameterizedType)) {
                        throw new RuntimeException("集合类型必须指定泛型参数");
                    }

                    ParameterizedType parameterizedType = (ParameterizedType) genericType;
                    Type[] types = parameterizedType.getActualTypeArguments();
                    if (types.length != 1)
                        throw new RuntimeException("集合类型参数错误");
                    Class<?> typeClass = (Class<?>) types[0];

                    Stream<Object> stream = IntStream
                            .range(fieldCell.getStartCellIndex(), fieldCell.getEndCellIndex() + 1)
                            .mapToObj(cellIndex -> getValue(typeClass, sheet, rowIndex, cellIndex, fieldCell.getFormatter()));

                    // 使用 Class 判断而不是 instanceof
                    if (List.class.isAssignableFrom(typeClazz)) {
                        ObjectUtils.setField(obj, fieldCell.getField(), stream.collect(Collectors.toList()));
                    } else if (Set.class.isAssignableFrom(typeClazz)) {
                        ObjectUtils.setField(obj, fieldCell.getField(), stream.collect(Collectors.toSet()));
                    } else {
                        throw new RuntimeException("暂不支持该集合类型: " + typeClazz.getName());
                    }
                } else {
                    // 普通类型处理
                    if (!Objects.equals(fieldCell.getStartCellIndex(), fieldCell.getEndCellIndex()))
                        throw new RuntimeException("该数据类型不支持多列");
                    ObjectUtils.setField(obj, fieldCell.getField(),
                            getValue(typeClazz, sheet, rowIndex, fieldCell.getStartCellIndex(),
                                    fieldCell.getFormatter()));
                }
            }
            return obj;
        }).filter(ObjectUtils::isNotEmpty).collect(Collectors.toList());

    }

    /**
     * 获取表头信息
     * <p>
     * 解析 Excel 表头结构，支持多行表头和合并单元格的处理
     * </p>
     *
     * @param sheet    Excel工作表
     * @param isSingle 是否将多行表头合并为单行
     * @param doExcel  Excel读取监听器
     * @return 表头单元格列表
     */
    public static List<HeaderCell> getHeaders(Sheet sheet, boolean isSingle, IExcelReaderListener doExcel) {
        if (sheet == null)
            throw new RuntimeException("表格数据不能为空");

        // 获取表头行范围
        int startHeaderNumber;
        int endHeaderNumber;
        if (doExcel != null) {
            startHeaderNumber = doExcel.startHeaderNumber(sheet) - 1;
            endHeaderNumber = doExcel.endHeaderNumber(sheet);
        } else {
            startHeaderNumber = 0;
            endHeaderNumber = 0;
        }
        if (startHeaderNumber < 0)
            throw new RuntimeException("表头开始行数不正确");
        if (endHeaderNumber < 0)
            throw new RuntimeException("表头行数不能小于0");

        // 获取合并单元格信息
        List<CellRangeAddress> mergedRegionList = sheet.getMergedRegions().stream()
                .filter(it -> it.getFirstRow() <= endHeaderNumber).collect(Collectors.toList());
        List<HeaderCell> headerCellList = getHeaderCellList(sheet, startHeaderNumber, endHeaderNumber,
                mergedRegionList);

        if (!isSingle)
            return headerCellList;

        // 合并多行表头为单行
        List<HeaderCell> finalHeaderCellList;
        if (endHeaderNumber - startHeaderNumber > 1) {
            int startCellIndex = headerCellList.stream().mapToInt(HeaderCell::getStartCellIndex).min().orElse(0);
            int endCellIndex = headerCellList.stream().mapToInt(HeaderCell::getEndCellIndex).max().orElse(0);

            finalHeaderCellList = IntStream.range(startCellIndex, endCellIndex + 1).mapToObj(index -> {
                List<HeaderCell> list = headerCellList.stream()
                        .filter(it -> it.getStartCellIndex() <= index && it.getEndCellIndex() >= index)
                        .sorted(Comparator.comparing(HeaderCell::getRowIndex)).collect(Collectors.toList());

                if (CollectionUtils.isEmpty(list))
                    return null;
                if (list.size() == 1)
                    return list.get(0);
                else {
                    // 多行表头用连接符连接
                    String cellValue = list.stream()
                            .map(it -> it.getCellValue() == null ? "" : it.getCellValue().toString())
                            .collect(Collectors.joining(HEADER_SEPARATOR));
                    HeaderCell last = list.get(list.size() - 1);
                    last.setCellValue(cellValue);
                    return last;
                }
            }).collect(Collectors.toList());
        } else
            finalHeaderCellList = headerCellList;
        return finalHeaderCellList;
    }

    /**
     * 读取数据为Map格式
     * <p>
     * 根据提供的表头信息，将表格数据读取为 Map 列表，
     * 其中 key 为表头名称，value 为对应单元格的值
     * </p>
     *
     * @param sheet          Excel工作表
     * @param headerCellList 表头单元格列表
     * @param doExcel        Excel读取监听器
     * @return Map格式的数据列表
     */
    public static List<Map<String, Object>> doMap(Sheet sheet, List<HeaderCell> headerCellList,
                                                  IExcelReaderListener doExcel) {
        if (sheet == null)
            throw new RuntimeException("表格数据不能为空");
        if (CollectionUtils.isEmpty(headerCellList))
            throw new RuntimeException("表头数据不能为空");

        // 按列索引排序表头
        headerCellList
                .sort(Comparator.comparing(HeaderCell::getStartCellIndex).thenComparing(HeaderCell::getEndCellIndex));
        int headerNumber;
        if (doExcel != null) {
            headerNumber = doExcel.endHeaderNumber(sheet);
        } else {
            headerNumber = 0;
        }
        if (headerNumber < 0)
            throw new RuntimeException("表头行数不能小于0");

        // 逐行读取数据到Map
        return IntStream.range(headerNumber, sheet.getLastRowNum() + 1).mapToObj(rowIndex -> {
            Row row = sheet.getRow(rowIndex);
            if (doExcel != null && doExcel.isFooter(row))
                return null;

            Map<String, Object> map = new LinkedHashMap<>();
            for (HeaderCell headerCell : headerCellList) {
                Object obj;
                if (headerCell.getStartCellIndex() == headerCell.getEndCellIndex()) {
                    // 单列数据
                    obj = getCellValueWithMergedRegion(sheet, rowIndex, headerCell.getStartCellIndex());
                } else {
                    // 多列数据，返回列表
                    obj = IntStream.range(headerCell.getStartCellIndex(), headerCell.getEndCellIndex() + 1)
                            .mapToObj(cellIndex -> getCellValueWithMergedRegion(sheet, rowIndex, cellIndex)).collect(Collectors.toList());
                }
                map.put(headerCell.getCellValue().toString(), obj);
            }
            return map;
        }).collect(Collectors.toList());
    }

    /**
     * 读取数据为Map格式（自动解析表头）
     * <p>
     * 自动解析表头信息，并将表格数据读取为 Map 列表
     * </p>
     *
     * @param sheet   Excel工作表
     * @param doExcel Excel读取监听器
     * @return Map格式的数据列表
     */
    public static List<Map<String, Object>> doMap(Sheet sheet, IExcelReaderListener doExcel) {
        if (sheet == null)
            throw new RuntimeException("表格数据不能为空");

        // 自动获取表头信息
        List<HeaderCell> headerCellList = getHeaders(sheet, true, doExcel);
        if (CollectionUtils.isEmpty(headerCellList))
            throw new RuntimeException("表头数据不能为空");

        headerCellList
                .sort(Comparator.comparing(HeaderCell::getStartCellIndex).thenComparing(HeaderCell::getEndCellIndex));
        int headerNumber;
        if (doExcel != null) {
            headerNumber = doExcel.endHeaderNumber(sheet);
        } else {
            headerNumber = 0;
        }
        if (headerNumber < 0)
            throw new RuntimeException("表头行数不能小于0");

        return IntStream.range(headerNumber, sheet.getLastRowNum() + 1).mapToObj(rowIndex -> {
            Row row = sheet.getRow(rowIndex);
            if (doExcel != null && doExcel.isFooter(row))
                return null;

            Map<String, Object> map = new LinkedHashMap<>();
            for (HeaderCell headerCell : headerCellList) {
                Object obj;
                if (headerCell.getStartCellIndex() == headerCell.getEndCellIndex()) {
                    obj = getCellValueWithMergedRegion(sheet, rowIndex, headerCell.getStartCellIndex());
                } else {
                    obj = IntStream.range(headerCell.getStartCellIndex(), headerCell.getEndCellIndex() + 1)
                            .mapToObj(cellIndex -> getCellValueWithMergedRegion(sheet, rowIndex, cellIndex)).collect(Collectors.toList());
                }
                map.put(headerCell.getCellValue().toString(), obj);
            }
            return map;
        }).collect(Collectors.toList());
    }

    /**
     * 解析字段与单元格的映射关系
     * <p>
     * 根据 @ExcelReader 注解配置，建立字段与Excel单元格的对应关系
     * </p>
     *
     * @param clazz             目标类
     * @param sheet             Excel工作表
     * @param startHeaderNumber 表头开始行
     * @param endHeaderNumber   表头结束行
     * @return 字段单元格映射列表
     */
    private static List<FieldCell> getFieldCells(Class<?> clazz, Sheet sheet, Integer startHeaderNumber,
                                                 Integer endHeaderNumber) {
        Field[] fields = clazz.getDeclaredFields();
        return Arrays.stream(fields).map(field -> {
            FieldCell fieldCell = new FieldCell();
            fieldCell.setField(field);
            if (!field.isAnnotationPresent(ExcelReader.class))
                return fieldCell;

            ExcelReader annotation = field.getAnnotation(ExcelReader.class);
            if (annotation == null)
                return fieldCell;

            // 设置格式化器
            if (annotation.formatter() != DefaultExcelReaderFormatter.class) {
                try {
                    fieldCell.setFormatter(annotation.formatter().getDeclaredConstructor().newInstance());
                } catch (Exception e) {
                    throw new RuntimeException("创建格式化器失败: " + e.getMessage(), e);
                }
            }

            String cellValue = annotation.value();
            int cellIndex = annotation.index();
            String[] cellNames = annotation.name();

            if (StringUtils.isNotEmpty(cellValue)) {
                // 坐标方式定位（如 "A1"）
                Matcher matcher = COORDINATE_PATTERN.matcher(cellValue);
                if (matcher.find()) {
                    String[] parts = {matcher.group(1), matcher.group(2)};
                    fieldCell.setRowIndex(Integer.parseInt(parts[1]) - 1);
                    fieldCell.setStartCellIndex(ObjectUtils.convertToNumber(parts[0]) - 1);
                    fieldCell.setEndCellIndex(ObjectUtils.convertToNumber(parts[0]) - 1);
                }
            } else if (cellIndex != -1) {
                // 索引方式定位
                fieldCell.setStartCellIndex(cellIndex);
                fieldCell.setEndCellIndex(cellIndex);
            } else if (ArrayUtils.isNotEmpty(cellNames)) {
                // 表头名称方式定位
                List<HeaderCell> headerCellList;
                List<CellRangeAddress> mergedRegionList;
                if (sheet != null) {
                    mergedRegionList = sheet.getMergedRegions().stream()
                            .filter(it -> it.getFirstRow() <= endHeaderNumber).collect(Collectors.toList());
                    headerCellList = getHeaderCellList(sheet, startHeaderNumber, endHeaderNumber, mergedRegionList);
                } else
                    headerCellList = null;

                if (headerCellList == null)
                    throw new RuntimeException("无法解析表头数据");

                // 处理多级表头匹配
                HeaderCell tmpHeaderCell = null;
                for (int i = 0; i < cellNames.length; i++) {
                    String name = cellNames[i];
                    HeaderCell finalTmpHeaderCell = tmpHeaderCell;
                    List<HeaderCell> list = headerCellList.stream().filter(it -> {
                        // 多表头的情况，且父表头已获取到数据时
                        if (finalTmpHeaderCell != null) {
                            // 判断子表头肯定在父表头的下一行
                            if (finalTmpHeaderCell.getRowIndex() >= it.getRowIndex())
                                return false;
                            // 判断子表头列下标需在父表头的内部
                            return finalTmpHeaderCell.getStartCellIndex() <= it.getStartCellIndex()
                                    && finalTmpHeaderCell.getEndCellIndex() >= it.getEndCellIndex()
                                    && Objects.equals(name, it.getCellValue());
                        }
                        return Objects.equals(name, it.getCellValue());
                    }).collect(Collectors.toList());

                    // 没有对应数据时
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
                            throw new RuntimeException("表头过于复杂，推荐使用 @ExcelReader(index = ?) 方式处理数据");
                        } else {
                            // 过滤掉有父表头的数据
                            list = list.stream()
                                    .filter(it -> headerCellList.stream()
                                            .filter(headerCell -> headerCell.getRowIndex() != it.getRowIndex()
                                                    && headerCell.getStartCellIndex() <= it.getStartCellIndex()
                                                    && headerCell.getEndCellIndex() >= it.getEndCellIndex())
                                            .count() == 1)
                                    .collect(Collectors.toList());
                            if (CollectionUtils.isEmpty(list))
                                // 表头数据重复
                                throw new RuntimeException("表头过于复杂，推荐使用 @ExcelReader(index = ?) 方式处理数据");
                            else if (list.size() == 1)
                                // 单条数据
                                tmpHeaderCell = list.get(0);
                            else
                                // 子表头有重复数据
                                throw new RuntimeException("表头过于复杂，推荐使用 @ExcelReader(index = ?) 方式处理数据");
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

    /**
     * 获取表头单元格列表
     * <p>
     * 解析指定范围内的表头单元格，处理合并单元格的情况
     * </p>
     *
     * @param sheet             Excel工作表
     * @param startHeaderNumber 表头开始行
     * @param endHeaderNumber   表头结束行
     * @param mergedRegionList  合并单元格列表
     * @return 表头单元格列表
     */
    private static List<HeaderCell> getHeaderCellList(Sheet sheet, Integer startHeaderNumber, Integer endHeaderNumber,
                                                      List<CellRangeAddress> mergedRegionList) {
        if (startHeaderNumber >= endHeaderNumber)
            return Collections.emptyList();

        return IntStream.range(startHeaderNumber, endHeaderNumber).mapToObj(rowIndex -> {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                return new ArrayList<HeaderCell>();
            }
            return IntStream.range(0, row.getLastCellNum()).mapToObj(cellIndex -> {
                Cell cell = row.getCell(cellIndex);
                Object cellValue = getCellValue(cell);

                // 检查是否在合并单元格范围内
                CellRangeAddress cellAddresses = mergedRegionList.stream()
                        .filter(range -> range.getFirstRow() <= rowIndex && range.getLastRow() >= rowIndex
                                && range.getFirstColumn() <= cellIndex && range.getLastColumn() >= cellIndex)
                        .findFirst().orElse(null);

                if (cellAddresses != null) {
                    // 只有合并单元格的第一个单元格才返回HeaderCell
                    if (Objects.equals(cellAddresses.getFirstRow(), rowIndex)
                            && Objects.equals(cellAddresses.getFirstColumn(), cellIndex))
                        return new HeaderCell(cellValue, rowIndex, cellIndex, cellAddresses.getLastColumn());
                    else
                        return null;
                } else
                    return new HeaderCell(cellValue, rowIndex, cellIndex, cellIndex);
            }).filter(it -> it != null && it.getCellValue() != null).collect(Collectors.toList());
        }).flatMap(Collection::stream).collect(Collectors.toList());
    }

    /**
     * 根据类型获取单元格值
     * <p>
     * 将单元格的值转换为指定的Java类型，支持自定义格式化器
     * </p>
     *
     * @param clazz     目标类型
     * @param cell      单元格
     * @param formatter 格式化器
     * @return 转换后的值
     * @throws RuntimeException 当不支持的数据类型时抛出
     */
    public static Object getValue(Class<?> clazz, Cell cell, IExcelReaderFormatter<?> formatter) {
        if (cell == null)
            return null;

        // 优先使用自定义格式化器
        if (formatter != null)
            return formatter.format(cell);

        // 字符串类型
        if (clazz.equals(String.class)) {
            Object object = getCellValue(cell);
            return object == null ? null : object.toString();
        }

        // 数字类型
        if (Number.class.isAssignableFrom(clazz)) {
            try {
                double value = cell.getNumericCellValue();
                // 检查整数类型是否包含小数部分
                if ((clazz.equals(Integer.class) || clazz.equals(Long.class))
                        && (String.valueOf(value).matches(".*\\.\\d*[1-9]+\\d*$")))
                    throw new RuntimeException("当前数据是浮点类型，无法转换为整数");

                BigDecimal cellValue = BigDecimal.valueOf(value);
                if (clazz.equals(Integer.class))
                    return cellValue.intValue();
                if (clazz.equals(Double.class))
                    return cellValue.doubleValue();
                if (clazz.equals(Float.class))
                    return cellValue.floatValue();
                if (clazz.equals(Long.class))
                    return cellValue.longValue();
                if (clazz.equals(BigDecimal.class))
                    return cellValue;
            } catch (IllegalStateException e) {
                // 尝试从字符串解析数字
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
                if (clazz.equals(BigDecimal.class))
                    return new BigDecimal(cellValue);
                return null;
            }
        }

        // 日期类型
        if (clazz.equals(Date.class))
            return new SimpleDateExcelReaderFormatter().format(cell);

        // 布尔类型
        if (clazz.equals(Boolean.class))
            return cell.getBooleanCellValue();

        throw new RuntimeException("暂不支持该数据类型: " + clazz.getName());
    }

    /**
     * 获取单元格的值，支持合并单元格
     * <p>
     * 当指定位置的单元格为空时，会检查是否在合并单元格范围内，
     * 如果是，则返回合并单元格的值
     * </p>
     *
     * @param clazz     目标类型
     * @param sheet     工作表
     * @param rowIndex  行索引
     * @param cellIndex 列索引
     * @param formatter 格式化器
     * @return 转换后的值
     */
    public static Object getValue(Class<?> clazz, Sheet sheet, int rowIndex, int cellIndex,
                                  IExcelReaderFormatter<?> formatter) {
        Row row = sheet.getRow(rowIndex);
        if (row == null)
            return null;

        Cell cell = row.getCell(cellIndex);
        Object value = null;
        if (cell != null) {
            value = getCellValue(cell);
        }

        // 如果当前单元格没有值，检查是否在合并单元格范围内
        if (value == null) {
            for (CellRangeAddress mergedRegion : sheet.getMergedRegions()) {
                if (mergedRegion.isInRange(rowIndex, cellIndex)) {
                    // 获取合并单元格的第一个单元格
                    Row firstRow = sheet.getRow(mergedRegion.getFirstRow());
                    if (firstRow != null) {
                        cell = firstRow.getCell(mergedRegion.getFirstColumn());
                        break;
                    }
                }
            }
        }

        return getValue(clazz, cell, formatter);
    }

    /**
     * 获取单元格原始值
     * <p>
     * 根据单元格类型返回对应的原始值，不进行类型转换
     * </p>
     *
     * @param cell 单元格
     * @return 单元格的原始值
     */
    public static Object getCellValue(Cell cell) {
        if (cell == null)
            return null;

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                // 判断是否为日期格式
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                // 公式类型，尝试获取计算后的值
                try {
                    return cell.getNumericCellValue();
                } catch (IllegalStateException e) {
                    return cell.getStringCellValue();
                }
            case BLANK:
                return null;
            case ERROR:
                return cell.getErrorCellValue();
            default:
                return null;
        }
    }

    /**
     * 获取单元格的值，考虑合并单元格的情况
     * <p>
     * 如果指定位置在合并单元格范围内，返回合并单元格的值
     * </p>
     *
     * @param sheet     工作表
     * @param rowIndex  行索引
     * @param cellIndex 列索引
     * @return 单元格的值
     */
    private static Object getCellValueWithMergedRegion(Sheet sheet, int rowIndex, int cellIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null)
            return null;

        Cell cell = row.getCell(cellIndex);
        Object value = getCellValue(cell);

        // 如果当前单元格没有值，检查是否在合并单元格范围内
        if (value == null) {
            for (CellRangeAddress mergedRegion : sheet.getMergedRegions()) {
                if (mergedRegion.isInRange(rowIndex, cellIndex)) {
                    // 获取合并单元格的第一个单元格的值
                    Row firstRow = sheet.getRow(mergedRegion.getFirstRow());
                    if (firstRow != null) {
                        Cell firstCell = firstRow.getCell(mergedRegion.getFirstColumn());
                        value = getCellValue(firstCell);
                        break;
                    }
                }
            }
        }

        return value;
    }

}
