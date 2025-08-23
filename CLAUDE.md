# Excel-IO Project Context

This document provides essential context for AI assistants working on the Excel-IO project.

## Project Overview

**Excel-IO** is a modern Java library for Excel file operations, designed to be intuitive and powerful. It was created as a simplified replacement for over-engineered Excel libraries, focusing on a fluent API design.

## Key Design Principles

1. **Fluent API**: Chainable method calls for readability
2. **XLSX Only**: Focused on modern Excel format (no legacy .xls support)
3. **Type Safety**: Automatic type detection and conversion
4. **Template-Based Styling**: Preserve complex Excel formatting with embedded templates
5. **Production-Ready**: Templates work in both development and JAR deployment
6. **Bulk Operations**: Efficient processing of large datasets
7. **Encoding Support**: Handle international characters properly

## Project Structure

```
src/main/java/com/gwill/io/excel/
├── ExcelIO.java                    # Main facade class - entry point for all operations
├── ExcelReader.java                # Unified reader (basic + type-safe reading)
├── ExcelWriter.java                # Basic Excel writing with fluent API
├── CategorizedWriter.java          # Hierarchical summary/detail reports
├── CategorizedTemplateWriter.java  # Template-based categorized reports
├── AlternatingRowsWriter.java      # Zebra-striped tables for readability
├── FormTemplateWriter.java         # Form-based business document generation
├── ExcelJavaDataType.java          # Enum for type conversion bridge
├── ExcelIOException.java           # Custom exception handling
├── util/
│   ├── ResourceUtil.java            # Production-ready template loading utility
│   └── FileCopier.java              # Memory-based operations to avoid POI conflicts
└── examples/                       # Comprehensive usage examples
    ├── BasicUsageExamples.java      # Basic operations + bulk data (150 records)
    ├── TypedReadingExamples.java    # Type-safe reading + large datasets (100+ records)
    ├── CategorizedReportExamples.java # Hierarchical reports + structured data
    ├── AlternatingRowExamples.java  # Zebra striping + loop/bulk patterns
    ├── RealWorldExample.java        # Complete e-commerce workflow
    └── ExampleRunner.java           # Utility to run all examples
```

## Core Classes Explained

### ExcelIO.java - Main Facade
- **Purpose**: Single entry point for all Excel operations
- **Key Methods**: 
  - `read()` → `ExcelReader` (handles both basic and typed reading)
  - `write()` → `ExcelWriter` 
  - `writeCategorized()` → `CategorizedWriter`
  - `alternatingRows()` → `AlternatingRowsWriter`
  - `formTemplate()` → `FormTemplateWriter`
  - `categorizedFromTemplate()` → `CategorizedTemplateWriter`

### ExcelReader.java - Unified Reading
- **Purpose**: Single reader class supporting both basic and type-safe reading
- **Key Features**:
  - Automatic detection of metadata for type conversion
  - Methods: `withMetadata()`, `withTypes()`, `asMaps()`, `asRows()`
  - Supports file paths, InputStreams, and multiple sheets

### ExcelJavaDataType.java - Type Bridge
- **Purpose**: Enum bridging Java class names to internal types
- **Key Method**: `fromClassName(String)` for metadata-driven type conversion
- **Supported Types**: STRING, INTEGER, DOUBLE, BIG_DECIMAL, LOCAL_DATE, BOOLEAN, etc.

## API Design Patterns

### Fluent API Example
```java
ExcelIO.write("output.xlsx")
    .sheet("Data")
    .header("Name", "Age")
    .row("John", 30)
    .save();
```

### Type-Safe Reading
```java
// Automatic detection based on metadata
List<Map<String, Object>> data = ExcelIO.read("data.xlsx")
    .withMetadata("metadata.xlsx")  // or .withTypes("name:java.lang.String", ...)
    .sheet("Sheet1")
    .asMaps();
```

### Template-Based Operations (Production-Ready)
```java
// Templates work in both development and production environments
import com.gwill.io.excel.util.ResourceUtil;
import java.io.InputStream;

// Resource-aware template loading
try (InputStream template = ResourceUtil.getInputStream("src/main/resources/examples/alternating_template.xlsx")) {
    ExcelIO.alternatingRows(template)
        .sheet("Report")
        .header("A", "B", "C")
        .row("1", "2", "3")
        .saveAs("output.xlsx");
}

// Bulk data operations for large datasets
List<Map<String, Object>> largeDataset = getLargeDataset(); // 1000+ records
try (InputStream template = ResourceUtil.getInputStream("src/main/resources/examples/alternating_template.xlsx")) {
    ExcelIO.alternatingRows(template)
        .sheet("Bulk Data", largeDataset)  // Pass entire collection
        .saveAs("bulk_report.xlsx");
}
```

## Recent Refactoring (Completed)

The project underwent major refactoring to solve POI style conflicts and improve the API:

### Critical Bug Fixes:
1. **POI Style Conflicts**: Implemented memory-based template processing using `FileCopier` utility
2. **Memory Operations**: All template writers now use byte arrays to avoid style ownership issues
3. **Stream Support**: Added `writeTo()` and `toByteArray()` methods for backend applications

### API Consolidations:
1. **Readers**: `ExcelReader` + `TypedExcelReader` → single `ExcelReader`
2. **Method Names**: `alternatingFromTemplate()` → `alternatingRows()`
3. **Class Names**: `AlternatingRowWriter` → `AlternatingRowsWriter`
4. **Template Writers**: `TemplateWriter` → `FormTemplateWriter` (specialized for business forms)

### New Features:
- **FormTemplateWriter**: 16-row business document structure (invoices, purchase orders, reports)
- **FileCopier Utility**: Memory-based operations to prevent POI conflicts
- **Backend Support**: Stream-based output for web applications
- **Template Preservation**: Styles are copied rather than modified directly

### API Improvements:
- Eliminated confusing `readTyped()` methods (now deprecated)
- Single `read()` method automatically handles both basic and typed reading
- `fromTemplate()` replaced with `formTemplate()` for clarity
- More descriptive method names throughout

## Dependencies

Minimal dependencies for better compatibility:

- **Apache POI 5.3.0**: Excel file processing
- **Lombok**: Reducing boilerplate (provided scope)
- **Java 22**: Modern Java features

### Test Dependencies
- **JUnit Jupiter 5.10.0**: Unit testing framework
- **AssertJ 3.24.2**: Fluent assertions

### Built-in Utilities
- **ResourceUtil**: Production-ready template loading (filesystem → JAR fallback)
- **Pre-built Templates**: Embedded in `src/main/resources/examples/`

**Note**: Removed dependency on legacy `com.gwill.common.excel` library - this project is its replacement.

## Build Commands

```bash
# Compile
mvn compile

# Package
mvn package

# Run all examples (recommended)
mvn exec:java -Dexec.mainClass="com.gwill.io.excel.examples.ExampleRunner"

# Run individual examples
mvn exec:java -Dexec.mainClass="com.gwill.io.excel.examples.BasicUsageExamples"
mvn exec:java -Dexec.mainClass="com.gwill.io.excel.examples.TypedReadingExamples"
mvn exec:java -Dexec.mainClass="com.gwill.io.excel.examples.CategorizedReportExamples"
mvn exec:java -Dexec.mainClass="com.gwill.io.excel.examples.AlternatingRowExamples"
mvn exec:java -Dexec.mainClass="com.gwill.io.excel.examples.RealWorldExample"

# Note: mvn test runs JUnit tests (when created), not examples
```

## Testing Strategy

Examples demonstrate real-world usage patterns:
- `BasicUsageExamples.java` - Core functionality + bulk operations (150 records)
- `TypedReadingExamples.java` - Type conversion + large dataset processing (100+ orders)  
- `CategorizedReportExamples.java` - Hierarchical reports + structured data
- `AlternatingRowExamples.java` - Styled tables + loop/bulk patterns
- `RealWorldExample.java` - Complete e-commerce workflow with production patterns
- `ExampleRunner.java` - Run all examples with timing and error handling

## Production Features

### Resource-Aware Template Loading
```java
// Works in both development (filesystem) and production (JAR)
try (InputStream template = ResourceUtil.getInputStream("src/main/resources/examples/alternating_template.xlsx")) {
    ExcelIO.alternatingRows(template)...
}
```

### Pre-built Templates
Location: `src/main/resources/examples/`
- `alternating_template.xlsx` - Zebra-striped tables
- `categorized_template.xlsx` - Categorized reports (header/category/detail)
- `form_template.xlsx` - Business forms (invoices, purchase orders, reports)

### Bulk Data Patterns
```java
// Loop-based for custom processing
for (SalesRecord sale : salesData) {
    writer.row(sale.getDate(), sale.getProduct(), sale.getAmount());
}

// Collection-based for performance
List<Map<String, Object>> data = getLargeDataset();
ExcelIO.alternatingRows(template)
    .sheet("Bulk Data", data)  // Pass entire collection
    .saveAs("output.xlsx");
```

## Common Operations

### Writing Excel Files (Production Patterns)
```java
// Basic writing with bulk data
List<Map<String, Object>> data = getLargeDataset();
ExcelIO.write("file.xlsx").sheet("Data", data).save();

// Categorized reports with structured data
Map<String, List<RegionData>> regionData = getRegionalData();
var writer = ExcelIO.writeCategorized("report.xlsx").sheet("Sales").header(...);
for (String region : regionData.keySet()) {
    writer.category(region, calculateTotals(regionData.get(region)));
    for (RegionData data : regionData.get(region)) {
        writer.detail("  " + data.getName(), data.getMetrics());
    }
}
writer.save();

// Template-based styling (resource-aware)
try (InputStream template = ResourceUtil.getInputStream("src/main/resources/examples/alternating_template.xlsx")) {
    ExcelIO.alternatingRows(template).sheet("Data").header(...).saveAs("output.xlsx");
}

// Form-based business documents (invoices, purchase orders, reports)
try (InputStream template = ResourceUtil.getInputStream("src/main/resources/examples/form_template.xlsx")) {
    ExcelIO.formTemplate(template)
        .setValue("invoiceNumber", "INV-001")
        .setValue("customerName", "Acme Corp")
        .setValue("invoiceDate", LocalDate.now())
        .lineItem("LAP001", "Dell Laptop", 2, 1500.00, 3000.00)
        .lineItem("MOU001", "Wireless Mouse", 5, 25.00, 125.00)
        .setValue("subtotal", 3125.00)
        .setValue("tax", 312.50)
        .setValue("grandTotal", 3437.50)
        .saveAs("generated_invoice.xlsx");
}
```

### Reading Excel Files
```java
// Basic reading
List<Map<String, Object>> data = ExcelIO.read("file.xlsx").asMaps();

// Type-safe reading with metadata
List<Map<String, Object>> data = ExcelIO.read("file.xlsx")
    .withMetadata("metadata.xlsx")  // or .withTypes(...)
    .asMaps();

// Process large datasets efficiently
List<Map<String, Object>> orders = ExcelIO.read("large_orders.xlsx")
    .withTypes("orderDate:java.time.LocalDate", "amount:java.math.BigDecimal")
    .asMaps();
    
// Type-safe business logic
BigDecimal totalRevenue = orders.stream()
    .map(order -> (BigDecimal) order.get("amount"))
    .reduce(BigDecimal.ZERO, BigDecimal::add);
```

## Encoding Considerations

All writers support encoding specification:
```java
.encoding("UTF-8")  // Default, supports international characters
.encoding("GBK")    // Chinese
.encoding("Shift_JIS")  // Japanese
```

## Error Handling

All operations throw `ExcelIOException` for Excel-related errors. The library includes comprehensive error messages and proper exception chaining.

## Future Enhancements

- Performance optimizations for very large files
- Additional template formats
- More data type support
- Streaming operations for memory efficiency

---

**Note**: This project represents a successful refactoring from an over-engineered Excel library to a clean, intuitive API that balances simplicity with powerful features.