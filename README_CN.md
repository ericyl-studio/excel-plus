# Excel Plus

基于 Apache POI 库的 Excel 读写工具，提供简单易用的注解配置方式，支持复杂 Excel 操作。

## 特性

### Excel 读取

- **单对象读取**：通过坐标精确定位读取单元格数据
- **列表读取**：支持通过索引或表头名称批量读取数据
- **多表头支持**：自动处理复杂的多级表头结构
- **合并单元格**：智能识别并处理合并单元格
- **自定义格式化**：支持自定义数据格式化器
- **类型自动转换**：自动转换常见数据类型

### Excel 写入

- **坐标写入**：精确控制数据写入位置
- **对象写入**：将对象数据映射到 Excel
- **列表写入**：批量写入列表数据，自动生成表头
- **分页写入**：支持大数据量分页写入，避免内存溢出
- **复杂表格**：支持多表头、表尾、合并单元格等复杂结构
- **样式配置**：支持单元格样式、边框、对齐方式等配置

## 快速开始

### 安装

#### Gradle

```gradle
dependencies {
    implementation('com.ericyl.excel:excel-plus:0.1.17')
}
```

#### Maven

```xml
<dependency>
    <groupId>com.ericyl.excel</groupId>
    <artifactId>excel-plus</artifactId>
    <version>0.1.16</version>
</dependency>
```

### 基本使用

#### 1. Excel 读取

##### 单对象读取

```java
// 定义数据模型
public class Report {
    @ExcelReader(value = "A1")  // 读取 A1 单元格
    private String title;

    @ExcelReader(value = "B2")  // 读取 B2 单元格
    private Double amount;

    // getter/setter...
}

// 读取数据
Workbook workbook = WorkbookFactory.create(new File("report.xlsx"));
Sheet sheet = workbook.getSheetAt(0);
Report report = ExcelReaderUtils.doIt(sheet, Report.class);
```

##### 列表读取（通过索引）

```java
public class User {
    @ExcelReader(index = 0)  // 第一列
    private String name;

    @ExcelReader(index = 1)  // 第二列
    private Integer age;

    @ExcelReader(index = 2, formatter = DateFormatter.class)  // 自定义格式化
    private Date birthDate;

    // getter/setter...
}

// 读取数据
List<User> users = ExcelReaderUtils.doList(sheet, User.class, null);
```

##### 列表读取（通过表头）

```java
public class Product {
    @ExcelReader(name = "产品名称")
    private String name;

    @ExcelReader(name = {"价格信息", "单价"})  // 多级表头
    private BigDecimal price;

    @ExcelReader(name = {"价格信息", "折扣"})
    private Double discount;

    // getter/setter...
}

// 自定义读取行为
IExcelReaderListener listener = new IExcelReaderListener() {
    @Override
    public int endHeaderNumber(Sheet sheet) {
        return 3;  // 表头结束行号是第3行
    }

    @Override
    public boolean isFooter(Row row) {
        // 判断是否为表尾（如合计行）
        Cell firstCell = row.getCell(0);
        return firstCell != null && "合计".equals(firstCell.getStringCellValue());
    }
};

List<Product> products = ExcelReaderUtils.doList(sheet, Product.class, listener);
```

#### 2. Excel 写入

##### 单对象写入

```java
public class Summary {
    @ExcelWriter(value = "A1")
    private String title = "月度报表";

    @ExcelWriter(value = "B2")
    private Date reportDate = new Date();

    @ExcelWriter(value = "C3", formatter = CurrencyFormatter.class)
    private BigDecimal total = new BigDecimal("10000.50");
}

// 写入数据
Workbook workbook = new XSSFWorkbook();
Sheet sheet = workbook.createSheet("Summary");
Summary summary = new Summary();
ExcelWriterUtils.obj2Excel(workbook, sheet, summary);
```

##### 列表写入

```java
public class Employee {
    @ExcelWriter(name = "员工姓名", index = 0, width = 4000)
    private String name;

    @ExcelWriter(name = "部门", index = 1, width = 3000)
    private String department;

    @ExcelWriter(name = "入职日期", index = 2, formatter = DateFormatter.class)
    private Date joinDate;

    @ExcelWriter(name = "薪资", index = 3,
                 horizontalAlignment = HorizontalAlignment.RIGHT,
                 border = @ExcelWriterBorder(
                     value = {BorderValue.ALL},
                     style = BorderStyle.THIN
                 ))
    private BigDecimal salary;
}

// 写入数据
List<Employee> employees = getEmployees();
ExcelWriterUtils.list2Excel(workbook, sheet, employees, Employee.class);

// 保存文件
String filePath = ExcelWriterUtils.toFile("export", workbook);
```

##### 分页写入（大数据量）

```java
// 实现数据分页获取
IExcelWriterListener<List<Order>> listener = new IExcelWriterListener<List<Order>>() {
    @Override
    public List<Order> doSomething(int pageNumber, int pageSize) {
        // 从数据库分页查询
        return orderService.findByPage(pageNumber, pageSize);
    }
};

// 分页写入，每页1000条，共100页
ExcelWriterUtils.list2Excel(workbook, sheet, 100, 1000, Order.class, listener);
```

##### 复杂表格写入

```java
// 构建复杂表格结构
ExcelTable table = new ExcelTable();

// 设置多级表头
List<List<ExcelColumn>> headers = new ArrayList<>();
// 第一行表头
List<ExcelColumn> header1 = Arrays.asList(
    new ExcelColumn("基本信息", 1, 3),  // 跨3列
    new ExcelColumn("成绩信息", 1, 4)   // 跨4列
);
// 第二行表头
List<ExcelColumn> header2 = Arrays.asList(
    new ExcelColumn("姓名"),
    new ExcelColumn("学号"),
    new ExcelColumn("班级"),
    new ExcelColumn("语文"),
    new ExcelColumn("数学"),
    new ExcelColumn("英语"),
    new ExcelColumn("总分")
);
headers.add(header1);
headers.add(header2);
table.setHeaders(headers);

// 设置数据内容
List<List<ExcelColumn>> data = getData();
table.setColumns(data);

// 写入表格
ExcelWriterUtils.table2Excel(workbook, sheet, table);
```

### 自定义格式化器

#### 读取格式化器

```java
public class DateExcelReaderFormatter implements IExcelReaderFormatter<Date> {
    private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

    @Override
    public Date format(Cell cell) {
        if (cell == null) return null;

        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getDateCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            try {
                return sdf.parse(cell.getStringCellValue());
            } catch (ParseException e) {
                throw new RuntimeException("日期格式错误", e);
            }
        }
        return null;
    }
}
```

#### 写入格式化器

```java
public class CurrencyExcelWriterFormatter implements IExcelWriterFormatter {
    @Override
    public Object format(Object data) {
        if (data instanceof Number) {
            return String.format("¥%.2f", ((Number) data).doubleValue());
        }
        return data;
    }
}
```

## API 文档

### ExcelReaderUtils 主要方法

| 方法                                                                       | 说明             |
| -------------------------------------------------------------------------- | ---------------- |
| `doIt(Sheet sheet, Class<T> clazz)`                                        | 读取单个对象数据 |
| `doList(Sheet sheet, Class<T> clazz, IExcelReaderListener listener)`       | 读取列表数据     |
| `doMap(Sheet sheet, IExcelReaderListener listener)`                        | 读取为 Map 格式  |
| `getHeaders(Sheet sheet, boolean isSingle, IExcelReaderListener listener)` | 获取表头信息     |

### ExcelWriterUtils 主要方法

| 方法                                                                                                                         | 说明         |
| ---------------------------------------------------------------------------------------------------------------------------- | ------------ |
| `xy(Workbook workbook, Sheet sheet, String xy, Object obj)`                                                                  | 坐标方式写入 |
| `obj2Excel(Workbook workbook, Sheet sheet, T obj)`                                                                           | 对象方式写入 |
| `list2Excel(Workbook workbook, Sheet sheet, List<T> list, Class<T> clazz)`                                                   | 列表方式写入 |
| `list2Excel(Workbook workbook, Sheet sheet, int page, int pageSize, Class<T> clazz, IExcelWriterListener<List<T>> listener)` | 分页方式写入 |
| `table2Excel(Workbook workbook, Sheet sheet, ExcelTable table)`                                                              | 复杂表格写入 |
| `toFile(String filePath, Workbook workbook)`                                                                                 | 保存到文件   |

## 注意事项

1. **内存管理**：处理大文件时建议使用分页读写功能
2. **类型转换**：自动类型转换仅支持常见类型，复杂类型需要自定义格式化器
3. **性能优化**：大数据量写入时，建议使用 SXSSFWorkbook 代替 XSSFWorkbook
4. **错误处理**：注意处理可能的运行时异常

## 示例项目

更多使用示例请参考 `src/test/java/com/ericyl/excel` 目录中的代码。

## 依赖库

- [Lombok](https://github.com/projectlombok/lombok) - 简化 Java 代码
- [Apache Commons](https://commons.apache.org) - 通用工具库
- [Apache POI](https://poi.apache.org) - Excel 操作核心库

## License

本项目采用 Apache License 2.0 协议，详见 [LICENSE](LICENSE) 文件。
