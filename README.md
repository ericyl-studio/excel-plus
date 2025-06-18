# Excel Plus

An annotation-based Excel reading and writing utility built on Apache POI, providing simple and easy-to-use configuration for complex Excel operations.

## Features

### Excel Reading

- **Single Object Reading**: Precisely read cell data through coordinate positioning
- **List Reading**: Batch read data through index or header name
- **Multi-Header Support**: Automatically handle complex multi-level header structures
- **Merged Cell Support**: Intelligently recognize and process merged cells
- **Custom Formatting**: Support for custom data formatters
- **Automatic Type Conversion**: Automatically convert common data types

### Excel Writing

- **Coordinate Writing**: Precisely control data writing position
- **Object Writing**: Map object data to Excel
- **List Writing**: Batch write list data with automatic header generation
- **Paginated Writing**: Support for large data pagination to avoid memory overflow
- **Complex Tables**: Support multi-headers, footers, merged cells and other complex structures
- **Style Configuration**: Support cell styles, borders, alignment and other configurations

## Quick Start

### Installation

#### Gradle

```gradle
dependencies {
    implementation('com.ericyl.excel:excel-plus:0.1.16')
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

### Basic Usage

#### 1. Excel Reading

##### Single Object Reading

```java
// Define data model
public class Report {
    @ExcelReader(value = "A1")  // Read cell A1
    private String title;

    @ExcelReader(value = "B2")  // Read cell B2
    private Double amount;

    // getter/setter...
}

// Read data
Workbook workbook = WorkbookFactory.create(new File("report.xlsx"));
Sheet sheet = workbook.getSheetAt(0);
Report report = ExcelReaderUtils.doIt(sheet, Report.class);
```

##### List Reading (By Index)

```java
public class User {
    @ExcelReader(index = 0)  // First column
    private String name;

    @ExcelReader(index = 1)  // Second column
    private Integer age;

    @ExcelReader(index = 2, formatter = DateFormatter.class)  // Custom formatting
    private Date birthDate;

    // getter/setter...
}

// Read data
List<User> users = ExcelReaderUtils.doList(sheet, User.class, null);
```

##### List Reading (By Header)

```java
public class Product {
    @ExcelReader(name = "Product Name")
    private String name;

    @ExcelReader(name = {"Price Info", "Unit Price"})  // Multi-level header
    private BigDecimal price;

    @ExcelReader(name = {"Price Info", "Discount"})
    private Double discount;

    // getter/setter...
}

// Custom reading behavior
IExcelReaderListener listener = new IExcelReaderListener() {
    @Override
    public int endHeaderNumber(Sheet sheet) {
        return 3;  // Header ends at row 3
    }

    @Override
    public boolean isFooter(Row row) {
        // Determine if it's a footer (like a total row)
        Cell firstCell = row.getCell(0);
        return firstCell != null && "Total".equals(firstCell.getStringCellValue());
    }
};

List<Product> products = ExcelReaderUtils.doList(sheet, Product.class, listener);
```

#### 2. Excel Writing

##### Single Object Writing

```java
public class Summary {
    @ExcelWriter(value = "A1")
    private String title = "Monthly Report";

    @ExcelWriter(value = "B2")
    private Date reportDate = new Date();

    @ExcelWriter(value = "C3", formatter = CurrencyFormatter.class)
    private BigDecimal total = new BigDecimal("10000.50");
}

// Write data
Workbook workbook = new XSSFWorkbook();
Sheet sheet = workbook.createSheet("Summary");
Summary summary = new Summary();
ExcelWriterUtils.obj2Excel(workbook, sheet, summary);
```

##### List Writing

```java
public class Employee {
    @ExcelWriter(name = "Employee Name", index = 0, width = 4000)
    private String name;

    @ExcelWriter(name = "Department", index = 1, width = 3000)
    private String department;

    @ExcelWriter(name = "Join Date", index = 2, formatter = DateFormatter.class)
    private Date joinDate;

    @ExcelWriter(name = "Salary", index = 3,
                 horizontalAlignment = HorizontalAlignment.RIGHT,
                 border = @ExcelWriterBorder(
                     value = {BorderValue.ALL},
                     style = BorderStyle.THIN
                 ))
    private BigDecimal salary;
}

// Write data
List<Employee> employees = getEmployees();
ExcelWriterUtils.list2Excel(workbook, sheet, employees, Employee.class);

// Save file
String filePath = ExcelWriterUtils.toFile("export", workbook);
```

##### Paginated Writing (Large Datasets)

```java
// Implement data pagination retrieval
IExcelWriterListener<List<Order>> listener = new IExcelWriterListener<List<Order>>() {
    @Override
    public List<Order> doSomething(int pageNumber, int pageSize) {
        // Query database with pagination
        return orderService.findByPage(pageNumber, pageSize);
    }
};

// Paginated writing, 1000 records per page, 100 pages total
ExcelWriterUtils.list2Excel(workbook, sheet, 100, 1000, Order.class, listener);
```

##### Complex Table Writing

```java
// Build complex table structure
ExcelTable table = new ExcelTable();

// Set multi-level headers
List<List<ExcelColumn>> headers = new ArrayList<>();
// First row header
List<ExcelColumn> header1 = Arrays.asList(
    new ExcelColumn("Basic Info", 1, 3),  // Span 3 columns
    new ExcelColumn("Score Info", 1, 4)   // Span 4 columns
);
// Second row header
List<ExcelColumn> header2 = Arrays.asList(
    new ExcelColumn("Name"),
    new ExcelColumn("Student ID"),
    new ExcelColumn("Class"),
    new ExcelColumn("Chinese"),
    new ExcelColumn("Math"),
    new ExcelColumn("English"),
    new ExcelColumn("Total")
);
headers.add(header1);
headers.add(header2);
table.setHeaders(headers);

// Set data content
List<List<ExcelColumn>> data = getData();
table.setColumns(data);

// Write table
ExcelWriterUtils.table2Excel(workbook, sheet, table);
```

### Custom Formatters

#### Read Formatter

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
                throw new RuntimeException("Date format error", e);
            }
        }
        return null;
    }
}
```

#### Write Formatter

```java
public class CurrencyExcelWriterFormatter implements IExcelWriterFormatter {
    @Override
    public Object format(Object data) {
        if (data instanceof Number) {
            return String.format("$%.2f", ((Number) data).doubleValue());
        }
        return data;
    }
}
```

## API Documentation

### ExcelReaderUtils Main Methods

| Method                                                                     | Description             |
| -------------------------------------------------------------------------- | ----------------------- |
| `doIt(Sheet sheet, Class<T> clazz)`                                        | Read single object data |
| `doList(Sheet sheet, Class<T> clazz, IExcelReaderListener listener)`       | Read list data          |
| `doMap(Sheet sheet, IExcelReaderListener listener)`                        | Read as Map format      |
| `getHeaders(Sheet sheet, boolean isSingle, IExcelReaderListener listener)` | Get header information  |

### ExcelWriterUtils Main Methods

| Method                                                                                                                       | Description          |
| ---------------------------------------------------------------------------------------------------------------------------- | -------------------- |
| `xy(Workbook workbook, Sheet sheet, String xy, Object obj)`                                                                  | Write by coordinates |
| `obj2Excel(Workbook workbook, Sheet sheet, T obj)`                                                                           | Write by object      |
| `list2Excel(Workbook workbook, Sheet sheet, List<T> list, Class<T> clazz)`                                                   | Write by list        |
| `list2Excel(Workbook workbook, Sheet sheet, int page, int pageSize, Class<T> clazz, IExcelWriterListener<List<T>> listener)` | Write by pagination  |
| `table2Excel(Workbook workbook, Sheet sheet, ExcelTable table)`                                                              | Write complex table  |
| `toFile(String filePath, Workbook workbook)`                                                                                 | Save to file         |

## Best Practices

1. **Memory Management**: Use pagination reading/writing functionality when processing large files
2. **Type Conversion**: Automatic type conversion only supports common types; complex types require custom formatters
3. **Performance Optimization**: For large data writes, use SXSSFWorkbook instead of XSSFWorkbook
4. **Error Handling**: Pay attention to possible runtime exceptions

## Language Support

- [English Documentation](README.md) (Current)
- [中文文档](README_CN.md)

## Example Code

For more usage examples, please refer to the code in the `com.ericyl.excel.example` package.

## Dependencies

- [Lombok](https://github.com/projectlombok/lombok) - Java code simplification
- [Apache Commons](https://commons.apache.org) - Common utility library
- [Apache POI](https://poi.apache.org) - Excel operation core library

## License

This project is licensed under the Apache License 2.0. See the [LICENSE](LICENSE) file for details.
