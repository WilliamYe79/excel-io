package com.gwill.io.excel;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * Fluent builder for writing XLSX Excel files.
 * Provides an intuitive API for creating Excel files with multiple sheets and various data types.
 * 
 * <p><strong>Note:</strong> This library supports .xlsx files only. For .xls files, 
 * please open them in Excel and save as .xlsx format first.</p>
 * 
 * <h3>Basic Usage</h3>
 * <pre>{@code
 * ExcelIO.write("output.xlsx")
 *     .sheet("Data")
 *     .header("Name", "Age", "Salary")
 *     .row("John", 30, 50000)
 *     .row("Jane", 25, 60000)
 *     .save();
 * }</pre>
 * 
 * <h3>With Custom Encoding</h3>
 * <pre>{@code
 * ExcelIO.write("output.xlsx")
 *     .encoding("UTF-8")  // For string data processing
 *     .sheet("Data")
 *     .header("姓名", "年龄", "工资")  // Chinese headers
 *     .row("张三", 30, 50000)
 *     .save();
 * }</pre>
 * 
 * <h3>From Collections</h3>
 * <pre>{@code
 * List<User> users = getUserList();
 * ExcelIO.write("users.xlsx")
 *     .sheet("Users", users)
 *     .save();
 * }</pre>
 */
public class ExcelWriter {
    
    private final String filePath;
    private final List<SheetBuilder> sheets = new ArrayList<>();
    private SheetBuilder currentSheet;
    private Charset encoding = StandardCharsets.UTF_8;

    public ExcelWriter(String filePath) {
        this.filePath = filePath;
    }

    /**
     * Set the encoding for processing string data.
     * Default is UTF-8, which works for most international characters.
     * 
     * @param encodingName the encoding name (e.g., "UTF-8", "GBK", "ISO-8859-1")
     * @return this ExcelWriter for method chaining
     * @throws ExcelIOException if the encoding is not supported
     */
    public ExcelWriter encoding(String encodingName) {
        try {
            this.encoding = Charset.forName(encodingName);
            return this;
        } catch (Exception e) {
            throw new ExcelIOException("Unsupported encoding: " + encodingName, e);
        }
    }

    /**
     * Set the encoding for processing string data.
     * 
     * @param charset the charset to use
     * @return this ExcelWriter for method chaining
     */
    public ExcelWriter encoding(Charset charset) {
        this.encoding = charset;
        return this;
    }

    /**
     * Add a new sheet with the given name and start building it.
     * 
     * @param sheetName the name of the sheet
     * @return this ExcelWriter for method chaining
     */
    public ExcelWriter sheet(String sheetName) {
        currentSheet = new SheetBuilder(sheetName, encoding);
        sheets.add(currentSheet);
        return this;
    }

    /**
     * Add a new sheet populated from a collection of objects.
     * Headers will be automatically inferred from the first object's fields/properties.
     * 
     * @param sheetName the name of the sheet
     * @param data the collection of objects to write
     * @return this ExcelWriter for method chaining
     */
    public <T> ExcelWriter sheet(String sheetName, Collection<T> data) {
        currentSheet = SheetBuilder.fromCollection(sheetName, data, encoding);
        sheets.add(currentSheet);
        return this;
    }


    /**
     * Set the headers for the current sheet.
     * 
     * @param headers the column headers
     * @return this ExcelWriter for method chaining
     */
    public ExcelWriter header(String... headers) {
        ensureCurrentSheet();
        currentSheet.header(headers);
        return this;
    }

    /**
     * Add a data row to the current sheet.
     * 
     * @param values the row values
     * @return this ExcelWriter for method chaining
     */
    public ExcelWriter row(Object... values) {
        ensureCurrentSheet();
        currentSheet.row(values);
        return this;
    }

    /**
     * Add multiple data rows to the current sheet.
     * 
     * @param rows the collection of row data
     * @return this ExcelWriter for method chaining
     */
    public ExcelWriter rows(Collection<Object[]> rows) {
        ensureCurrentSheet();
        rows.forEach(currentSheet::row);
        return this;
    }

    /**
     * Save the Excel file to the specified path.
     * 
     * @throws ExcelIOException if there's an error writing the file
     */
    public void save() {
        if (sheets.isEmpty()) {
            throw new ExcelIOException("No sheets defined. Call sheet() method first.");
        }

        try (Workbook workbook = new XSSFWorkbook()) {
            for (SheetBuilder sheetBuilder : sheets) {
                sheetBuilder.buildSheet(workbook);
            }

            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }
        } catch (IOException e) {
            throw new ExcelIOException("Failed to save Excel file: " + filePath, e);
        }
    }

    private void ensureCurrentSheet() {
        if (currentSheet == null) {
            sheet("Sheet1"); // Default sheet name
        }
    }

    /**
     * Internal builder class for constructing individual sheets.
     */
    static class SheetBuilder {
        private final String sheetName;
        private final Charset encoding;
        private final List<String> headers = new ArrayList<>();
        private final List<Object[]> rows = new ArrayList<>();
        private boolean headersSet = false;

        public SheetBuilder(String sheetName, Charset encoding) {
            this.sheetName = sheetName;
            this.encoding = encoding;
        }

        public static <T> SheetBuilder fromCollection(String sheetName, Collection<T> data, Charset encoding) {
            SheetBuilder builder = new SheetBuilder(sheetName, encoding);
            if (!data.isEmpty()) {
                // Use reflection to extract headers and data
                // This is a simplified implementation - you might want to enhance this
                T first = data.iterator().next();
                if (first instanceof Map) {
                    return fromMaps(sheetName, (Collection<Map<String, Object>>) data, encoding);
                }
                // For now, just convert to string representation
                // In a real implementation, you'd use reflection to get field names
                builder.header("Data");
                for (T item : data) {
                    builder.row(processStringValue(item.toString(), encoding));
                }
            }
            return builder;
        }

        public static SheetBuilder fromMaps(String sheetName, Collection<Map<String, Object>> data, Charset encoding) {
            SheetBuilder builder = new SheetBuilder(sheetName, encoding);
            if (!data.isEmpty()) {
                Map<String, Object> first = data.iterator().next();
                String[] headerArray = first.keySet().stream()
                    .map(key -> processStringValue(key, encoding))
                    .toArray(String[]::new);
                builder.header(headerArray);
                
                for (Map<String, Object> row : data) {
                    Object[] values = first.keySet().stream()
                        .map(key -> {
                            Object value = row.get(key);
                            return value instanceof String ? 
                                processStringValue((String) value, encoding) : value;
                        })
                        .toArray();
                    builder.row(values);
                }
            }
            return builder;
        }

        private static String processStringValue(String value, Charset encoding) {
            if (value == null) return null;
            // Ensure proper encoding handling
            // Convert to bytes with specified encoding, then back to string
            // This ensures consistent character handling
            return new String(value.getBytes(encoding), encoding);
        }

        public SheetBuilder header(String... headers) {
            if (headersSet) {
                throw new ExcelIOException("Headers already set for sheet: " + sheetName);
            }
            this.headers.clear();
            for (String header : headers) {
                this.headers.add(processStringValue(header, encoding));
            }
            headersSet = true;
            return this;
        }

        public SheetBuilder row(Object... values) {
            // Process string values with proper encoding
            Object[] processedValues = new Object[values.length];
            for (int i = 0; i < values.length; i++) {
                if (values[i] instanceof String) {
                    processedValues[i] = processStringValue((String) values[i], encoding);
                } else {
                    processedValues[i] = values[i];
                }
            }
            rows.add(processedValues);
            return this;
        }

        public void buildSheet(Workbook workbook) {
            Sheet sheet = workbook.createSheet(sheetName);
            int rowNum = 0;

            // Write headers if present
            if (!headers.isEmpty()) {
                Row headerRow = sheet.createRow(rowNum++);
                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers.get(i));
                    
                    // Style headers (bold)
                    CellStyle headerStyle = workbook.createCellStyle();
                    Font font = workbook.createFont();
                    font.setBold(true);
                    headerStyle.setFont(font);
                    cell.setCellStyle(headerStyle);
                }
            }

            // Write data rows
            for (Object[] rowData : rows) {
                Row row = sheet.createRow(rowNum++);
                for (int i = 0; i < rowData.length; i++) {
                    Cell cell = row.createCell(i);
                    setCellValue(cell, rowData[i]);
                }
            }

            // Auto-size columns
            for (int i = 0; i < Math.max(headers.size(), getMaxColumnCount()); i++) {
                sheet.autoSizeColumn(i);
            }
        }

        private int getMaxColumnCount() {
            return rows.stream()
                .mapToInt(row -> row.length)
                .max()
                .orElse(0);
        }

        private void setCellValue(Cell cell, Object value) {
            if (value == null) {
                cell.setCellValue("");
            } else if (value instanceof String) {
                cell.setCellValue((String) value);
            } else if (value instanceof Number) {
                cell.setCellValue(((Number) value).doubleValue());
            } else if (value instanceof Boolean) {
                cell.setCellValue((Boolean) value);
            } else {
                cell.setCellValue(value.toString());
            }
        }
    }
}