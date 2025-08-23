# Excel-IO

A modern, intuitive Java library for Excel file operations. Built with a fluent API design that makes reading and writing Excel files simple and powerful.

## Features

- **📖 Intuitive Reading**: Read Excel files with automatic type detection
- **✍️ Fluent Writing**: Create Excel files with a clean, chainable API  
- **🔧 Type-Safe Operations**: Specify data types for automatic conversion
- **🎨 Template-Based Styling**: Preserve complex Excel formatting using templates
- **📊 Categorized Reports**: Generate hierarchical summary/detail reports
- **🦓 Alternating Row Styles**: Create readable tables with zebra striping
- **📋 Form Templates**: Generate business documents (invoices, purchase orders, reports)
- **🌍 Encoding Support**: Handle international characters with UTF-8 and other encodings
- **⚡ XLSX Only**: Focused on modern Excel format for better performance
- **💾 Backend Support**: Stream-based output for web applications

## Quick Start

### Maven Dependency
```xml
<dependency>
    <groupId>com.gwill.io.excel</groupId>
    <artifactId>excel-io</artifactId>
    <version>1.0.0</version>
</dependency>
```

### Basic Writing
```java
// Create a simple Excel file
ExcelIO.write("employees.xlsx")
    .sheet("Staff")
    .header("Name", "Age", "Department", "Salary")
    .row("John Doe", 30, "Engineering", 75000)
    .row("Jane Smith", 28, "Marketing", 68000)
    .row("Mike Wilson", 35, "Sales", 62000)
    .save();
```

### Basic Reading
```java
// Read an Excel file
List<Map<String, Object>> employees = ExcelIO.read("employees.xlsx")
    .sheet("Staff")
    .asMaps();

for (Map<String, Object> emp : employees) {
    String name = (String) emp.get("Name");
    Long age = (Long) emp.get("Age");
    System.out.println(name + " is " + age + " years old");
}
```

## Advanced Features

### Type-Safe Reading

For precise data type handling, specify column types:

```java
// Using inline type definitions
List<Map<String, Object>> data = ExcelIO.read("financial_data.xlsx")
    .withTypes("Date:java.time.LocalDate", 
               "Amount:java.math.BigDecimal",
               "Processed:java.lang.Boolean")
    .sheet("Transactions")
    .asMaps();

// Using metadata file
List<Map<String, Object>> data = ExcelIO.read("sales_data.xlsx")
    .withMetadata("sales_metadata.xlsx")  // Contains column names and types
    .sheet("Sales")
    .asMaps();
```

### Encoding Support

Handle international characters properly:

```java
ExcelIO.write("chinese_data.xlsx")
    .encoding("UTF-8")  // Supports Chinese, Japanese, Korean, etc.
    .sheet("数据")
    .header("姓名", "年龄", "城市")
    .row("张三", 25, "北京")
    .row("李四", 30, "上海")
    .save();
```

### Categorized Reports

Create hierarchical reports with summary and detail rows:

```java
ExcelIO.writeCategorized("sales_report.xlsx")
    .sheet("Regional Sales")
    .header("Region/Product", "Units", "Revenue", "Profit")
    
    // Category: North America
    .category("North America", 1500, 150000, 45000)
    .detail("  Product A", 800, 80000, 24000)
    .detail("  Product B", 700, 70000, 21000)
    
    // Category: Europe  
    .category("Europe", 1200, 120000, 36000)
    .detail("  Product A", 600, 60000, 18000)
    .detail("  Product B", 600, 60000, 18000)
    
    .save();
```

### Template-Based Styling

Use Excel templates to preserve complex formatting:

```java
// First, create a template file with desired styles
ExcelIO.write("categorized_template.xlsx")
    .sheet("Template")
    .header("HEADER STYLE", "HEADER STYLE", "HEADER STYLE")
    .row("Category Style", "Category Style", "Category Style")  
    .row("Detail Style", "Detail Style", "Detail Style")
    .save();

// Then use the template for styled reports
ExcelIO.categorizedFromTemplate("categorized_template.xlsx")
    .sheet("Q1 Report")
    .header("Product", "Sales", "Profit")
    .category("Electronics", 50000, 15000)
    .detail("  Laptops", 30000, 9000)
    .detail("  Phones", 20000, 6000)
    .saveAs("styled_q1_report.xlsx");
```

### Alternating Row Styles (Zebra Striping)

Create readable tables with alternating row colors:

```java
// Create alternating row template
ExcelIO.write("table_template.xlsx")
    .sheet("Template")
    .header("Header Style", "Header Style", "Header Style")
    .row("Odd Row Style", "Odd Row Style", "Odd Row Style")
    .row("Even Row Style", "Even Row Style", "Even Row Style")
    .save();

// Apply alternating styles to your data
ExcelIO.alternatingRows("table_template.xlsx")
    .sheet("Employee Directory")
    .header("ID", "Name", "Department", "Salary")
    .row("001", "John Doe", "Engineering", 75000)
    .row("002", "Jane Smith", "Marketing", 68000)
    .row("003", "Mike Wilson", "Sales", 62000)
    // ... many more rows for better visual effect
    .saveAs("employee_directory.xlsx");
```

### Form-Based Business Documents

Create professional invoices, purchase orders, and reports using structured templates:

```java
// Create an invoice using a 12-row business form template
ExcelIO.formTemplate("form_template.xlsx")
    .setValue("invoiceNumber", "INV-001")
    .setValue("customerName", "Acme Corp")
    .setValue("customerAddress", "123 Business St, Commerce City, CA 90210")
    .setValue("invoiceDate", LocalDate.now())
    .setValue("dueDate", LocalDate.now().plusDays(30))
    
    // Add line items to the table section
    .lineItem("LAP001", "Dell Laptop XPS 13", 2, 1500.00, 3000.00)
    .lineItem("MOU001", "Wireless Mouse", 5, 25.00, 125.00)
    .lineItem("KEY001", "Mechanical Keyboard", 1, 150.00, 150.00)
    
    // Add footer calculations
    .setValue("subtotal", 3275.00)
    .setValue("tax", 327.50)
    .setValue("grandTotal", 3602.50)
    .setValue("paymentTerms", "Net 30 days")
    .setValue("notes", "Thank you for your business!")
    
    .saveAs("generated_invoice.xlsx");

// Form templates work great for backend applications too
ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
ExcelIO.formTemplate("form_template.xlsx")
    .setValue("poNumber", "PO-2024-001")
    .setValue("vendor", "Office Supplies Inc.")
    .lineItem("PEN-001", "Blue Pens", 100, 0.50, 50.00)
    .lineItem("PAP-001", "A4 Paper", 10, 5.00, 50.00)
    .setValue("total", 100.00)
    .writeTo(outputStream);
```

The form template structure follows a 16-row layout:
- **Rows 0-5**: Form header (company info, customer details, dates)
- **Row 6**: Table headers
- **Rows 7-8**: Line item style templates (odd/even alternating)
- **Row 9**: Separator row
- **Rows 10-15**: Form footer (totals, terms, signatures)

### Working with Collections

Easily convert Java objects to Excel:

```java
// From List of Maps
List<Map<String, Object>> products = getProductList();
ExcelIO.write("products.xlsx")
    .sheet("Inventory", products)
    .save();

// From any Collection
List<Employee> employees = getEmployees();
ExcelIO.write("employees.xlsx")
    .sheet("Staff", employees)
    .save();
```

## Supported Data Types

The library automatically handles these Java types:

- **Primitives**: `int`, `long`, `double`, `float`, `boolean`, `byte`, `short`
- **Wrapper Classes**: `Integer`, `Long`, `Double`, `Float`, `Boolean`, `Byte`, `Short`
- **Text**: `String`
- **Decimals**: `BigDecimal` (recommended for financial data)
- **Dates**: `LocalDate`, `LocalTime`, `LocalDateTime`

## Encoding Support

Excel-IO supports various character encodings:

- **UTF-8** (default) - Supports all international characters
- **GBK** - For simplified Chinese
- **ISO-8859-1** - For Western European languages
- **Shift_JIS** - For Japanese
- And any encoding supported by Java `Charset`

## Error Handling

The library uses `ExcelIOException` for all Excel-related errors:

```java
try {
    List<Map<String, Object>> data = ExcelIO.read("missing_file.xlsx")
        .asMaps();
} catch (ExcelIOException e) {
    System.err.println("Excel operation failed: " + e.getMessage());
    e.printStackTrace();
}
```

## Real-World Example

Here's a complete example similar to an e-commerce analysis system:

```java
// 1. Read raw sales data with proper types
List<Map<String, Object>> sales = ExcelIO.read("raw_sales_data.xlsx")
    .withTypes("OrderID:java.lang.String",
               "Date:java.time.LocalDate", 
               "Amount:java.math.BigDecimal",
               "CustomerType:java.lang.String")
    .sheet("Sales")
    .asMaps();

// 2. Process and generate categorized performance report
ExcelIO.writeCategorized("performance_report.xlsx")
    .sheet("Sales Performance")
    .header("Category/Product", "Orders", "Revenue", "Avg Order")
    .category("Electronics", 150, new BigDecimal("75000"), "$500")
    .detail("  Laptops", 80, new BigDecimal("48000"), "$600")
    .detail("  Phones", 70, new BigDecimal("27000"), "$386")
    .category("Books", 200, new BigDecimal("15000"), "$75")  
    .detail("  Fiction", 120, new BigDecimal("9000"), "$75")
    .detail("  Technical", 80, new BigDecimal("6000"), "$75")
    .save();

// 3. Generate detailed transaction log with alternating rows
ExcelIO.alternatingRows("transaction_template.xlsx")
    .encoding(StandardCharsets.UTF_8)
    .sheet("All Transactions")
    .header("Date", "Order ID", "Customer", "Amount", "Status")
    // Add hundreds of transactions here...
    .saveAs("transaction_log.xlsx");
```

## Performance Notes

- **XLSX Only**: This library only supports `.xlsx` files (not legacy `.xls`)
- **Memory Efficient**: Uses Apache POI's streaming capabilities where possible
- **Large Files**: For very large files (>100MB), consider processing in chunks
- **Templates**: Template-based operations are slightly slower but preserve complex formatting

## Dependencies

- **Apache POI 5.4.1** - Excel file processing
- **Lombok** - Reducing boilerplate code  
- **Java 22** - Modern Java features

## Migration from Legacy Systems

If you're migrating from complex Excel libraries:

**Before** (typical legacy approach):
```java
// Complex, verbose, error-prone
Workbook workbook = new XSSFWorkbook();
Sheet sheet = workbook.createSheet("Data");
Row headerRow = sheet.createRow(0);
headerRow.createCell(0).setCellValue("Name");
headerRow.createCell(1).setCellValue("Age");
Row dataRow = sheet.createRow(1);
dataRow.createCell(0).setCellValue("John");
dataRow.createCell(1).setCellValue(30);
FileOutputStream fileOut = new FileOutputStream("output.xlsx");
workbook.write(fileOut);
workbook.close();
fileOut.close();
```

**After** (Excel-IO):
```java
// Simple, readable, maintainable
ExcelIO.write("output.xlsx")
    .sheet("Data")
    .header("Name", "Age")
    .row("John", 30)
    .save();
```

## Examples

The library includes comprehensive examples demonstrating real-world patterns:

- **BasicUsageExamples** - Simple reading/writing + bulk operations (150 records)
- **TypedReadingExamples** - Type-safe conversion + large dataset processing (100+ orders)
- **CategorizedReportExamples** - Hierarchical reports + structured data processing
- **AlternatingRowExamples** - Zebra-striped tables + loop/bulk data patterns
- **FormTemplateExamples** - Business document generation (invoices, purchase orders, reports)
- **RealWorldExample** - Complete e-commerce analysis with production-ready patterns

### Pre-built Templates

The library includes production-ready templates in `src/main/resources/examples/`:

- **alternating_template.xlsx** - For zebra-striped tables
- **categorized_template.xlsx** - For categorized reports with header/category/detail styles  
- **form_template.xlsx** - For business forms (invoices, purchase orders, reports) with 16-row structure

These templates work automatically in both development and JAR environments.

### Running Examples

```bash
# Run all examples at once (recommended)
mvn exec:java -Dexec.mainClass="com.gwill.io.excel.examples.ExampleRunner"

# Run individual examples
mvn exec:java -Dexec.mainClass="com.gwill.io.excel.examples.BasicUsageExamples"
mvn exec:java -Dexec.mainClass="com.gwill.io.excel.examples.TypedReadingExamples"
mvn exec:java -Dexec.mainClass="com.gwill.io.excel.examples.CategorizedReportExamples"
mvn exec:java -Dexec.mainClass="com.gwill.io.excel.examples.AlternatingRowsExamples"
mvn exec:java -Dexec.mainClass="com.gwill.io.excel.examples.FormTemplateExamples"
mvn exec:java -Dexec.mainClass="com.gwill.io.excel.examples.RealWorldExample"

# Compile and package
mvn clean compile
mvn package
```

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- 📧 **Email**: [yeshengwei@gmail.com]
- 🐛 **Issues**: [GitHub Issues](https://github.com/your-username/excel-io/issues)
- 📖 **Documentation**: [Wiki](https://github.com/your-username/excel-io/wiki)

---

<div align="center">

**⭐ Would you give me a star if this project is helpful to you?**

**🔗 Share this project to more friends who need to read/write Excel spreadsheet files in their Java projects.**

Made with ❤️ by William YE of G-WILL Team

**Excel-IO** - Making Excel operations in Java intuitive and powerful. 🚀

</div>
