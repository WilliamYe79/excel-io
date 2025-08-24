package com.gwill.io.excel;

import com.gwill.io.excel.util.FileCopier;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * Form-based Excel writer that combines placeholder replacement with dynamic tabular data.
 * Perfect for business documents like invoices, purchase orders, and reports.
 *
 * <p>The template file follows a fixed 16-row structure:</p>
 * <ul>
 *   <li>Rows 0-5: Form header placeholders (customer info, dates, etc.)</li>
 *   <li>Row 6: Table headers</li>
 *   <li>Row 7: Odd row style template</li>
 *   <li>Row 8: Even row style template</li>
 *   <li>Row 9: Empty/separator</li>
 *   <li>Rows 10-15: Form footer placeholders (totals, signatures, etc.)</li>
 * </ul>
 *
 * <h2>Usage Example</h2>
 * <pre>{@code
 * ExcelIO.formTemplate("form_template.xlsx")
 *     .setValue("invoiceNumber", "INV-001")
 *     .setValue("customerName", "Acme Corp")
 *     .setValue("invoiceDate", LocalDate.now())
 *     .lineItem("LAP001", "Dell Laptop", 2, 1500.00, 3000.00)
 *     .lineItem("MOU001", "Wireless Mouse", 5, 25.00, 125.00)
 *     .setValue("subtotal", 3125.00)
 *     .setValue("tax", 312.50)
 *     .setValue("grandTotal", 3437.50)
 *     .saveAs("generated_invoice.xlsx");
 * }</pre>
 */
@RequiredArgsConstructor
public class FormTemplateWriter {

    public static final int FORM_HEADER_START = 0;
    public static final int FORM_HEADER_END = 5;
    public static final int TABLE_HEADER_ROW = 6;
    public static final int ODD_STYLE_ROW = 7;
    public static final int EVEN_STYLE_ROW = 8;
    public static final int TABLE_SEPARATOR_ROW = 9;
    public static final int FORM_FOOTER_START = 10;
    public static final int FORM_FOOTER_END = 15;

    private final String templatePath;
    private final InputStream templateStream;
    private final Map<String, Object> placeholderValues = new HashMap<>();
    private final List<Object[]> lineItems = new ArrayList<>();
    private Charset encoding = StandardCharsets.UTF_8;

    public FormTemplateWriter(String templatePath) {
        this.templatePath = templatePath;
        this.templateStream = null;
    }

    public FormTemplateWriter(InputStream templateStream) {
        this.templatePath = null;
        this.templateStream = templateStream;
    }

    /**
     * Set the encoding for processing string data.
     * Default is UTF-8, which works for most international characters.
     *
     * @param encodingName the encoding name (e.g., "UTF-8", "GBK", "ISO-8859-1")
     * @return this FormTemplateWriter for method chaining
     * @throws ExcelIOException if the encoding is not supported
     */
    public FormTemplateWriter encoding(String encodingName) {
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
     * @return this FormTemplateWriter for method chaining
     */
    public FormTemplateWriter encoding(Charset charset) {
        this.encoding = charset;
        return this;
    }

    /**
     * Set a placeholder value for form headers or footers.
     * Placeholders in template should be formatted as {{key}} or ${key}.
     *
     * @param key the placeholder key
     * @param value the value to replace the placeholder
     * @return this FormTemplateWriter for method chaining
     */
    public FormTemplateWriter setValue(String key, Object value) {
        placeholderValues.put(key, value);
        return this;
    }

    /**
     * Add multiple placeholder values at once.
     *
     * @param values the map of placeholder values
     * @return this FormTemplateWriter for method chaining
     */
    public FormTemplateWriter setValues(Map<String, Object> values) {
        placeholderValues.putAll(values);
        return this;
    }

    /**
     * Add a line item to the table section.
     * Line items will be inserted between the table header and footer with alternating styles.
     *
     * @param values the values for this line item
     * @return this FormTemplateWriter for method chaining
     */
    public FormTemplateWriter lineItem(Object... values) {
        Object[] processedValues = new Object[values.length];
        for (int i = 0; i < values.length; i++) {
            if (values[i] instanceof String) {
                processedValues[i] = processStringValue((String) values[i], encoding);
            } else {
                processedValues[i] = values[i];
            }
        }
        lineItems.add(processedValues);
        return this;
    }

    /**
     * Add multiple line items at once.
     *
     * @param items the collection of line item arrays
     * @return this FormTemplateWriter for method chaining
     */
    public FormTemplateWriter lineItems(Collection<Object[]> items) {
        items.forEach(this::lineItem);
        return this;
    }

    /**
     * Save the populated form template to the specified file path.
     * Uses in-memory byte array operations to avoid POI style conflicts.
     *
     * @param outputPath the path where the Excel file will be saved
     * @throws ExcelIOException if there's an error processing the template or saving the file
     */
    public void saveAs(String outputPath) {
        try {
            byte[] resultBytes = processTemplateInMemory();
            FileCopier.writeByteArrayToFile(resultBytes, outputPath);
        } catch (IOException e) {
            throw new ExcelIOException("Failed to process form template and save Excel file: " + outputPath, e);
        }
    }

    /**
     * Write the populated form template to the specified output stream.
     * This method is ideal for web applications and HTTP responses.
     *
     * @param outputStream the output stream to write to
     * @throws ExcelIOException if there's an error processing the template or writing to the stream
     */
    public void writeTo(OutputStream outputStream) {
        if (outputStream == null) {
            throw new IllegalArgumentException("Output stream cannot be null");
        }

        try {
            byte[] resultBytes = processTemplateInMemory();
            outputStream.write(resultBytes);
            outputStream.flush();
        } catch (IOException e) {
            throw new ExcelIOException("Failed to process form template and write to output stream", e);
        }
    }

    /**
     * Return the populated form template as a byte array.
     * This is useful for web applications or when you need to process the Excel data in memory.
     * Uses in-memory operations to avoid POI style conflicts.
     *
     * @return the populated Excel file as a byte array
     * @throws ExcelIOException if there's an error processing the template
     */
    public byte[] toByteArray() {
        try {
            return processTemplateInMemory();
        } catch (IOException e) {
            throw new ExcelIOException("Failed to process form template to byte array", e);
        }
    }

    /**
     * Process the template in memory to avoid POI style conflicts.
     * This method copies the template to memory, modifies it with data, and returns the result as bytes.
     * The template sheet is not deleted, but its content is modified with the provided data.
     */
    private byte[] processTemplateInMemory() throws IOException {
        // Step 1: Copy template to memory as byte array
        byte[] templateBytes;
        try (InputStream template = createTemplateStream()) {
            templateBytes = FileCopier.copyStreamToMemoryAsByteArray(template);
        }

        // Step 2: Create workbook from memory and populate with data
        try (InputStream templateByteStream = FileCopier.createInputStreamFromByteArray(templateBytes)) {
            try (Workbook workbook = new XSSFWorkbook(templateByteStream)) {

                // Use the first sheet and populate it with form data
                Sheet templateSheet = workbook.getSheetAt(0);
                populateFormSheet(templateSheet);

                // Return as byte array
                try (ByteArrayOutputStream byteOut = new ByteArrayOutputStream()) {
                    workbook.write(byteOut);
                    return byteOut.toByteArray();
                }
            }
        }
    }

    /**
     * Populate the form sheet with data following the 12-row template structure.
     * Implements the refined control logic to handle all edge cases properly.
     */
    private void populateFormSheet(Sheet templateSheet) {
        // Verify template has the required structure
        if (templateSheet.getLastRowNum() < FORM_FOOTER_END) {
            throw new ExcelIOException("Form template must have at least 16 rows (0-15) for the standard form structure");
        }

        int originalLastRow = templateSheet.getLastRowNum();

        // Step 1: Replace placeholders in header rows (0-5) and footer rows (10-15)
        replacePlaceholdersInRows(templateSheet, FORM_HEADER_START, FORM_HEADER_END);
        replacePlaceholdersInRows(templateSheet, FORM_FOOTER_START, FORM_FOOTER_END);

        // Step 2: Calculate line items space & shift footer if necessary
        int numLineItems = lineItems.size();
        int templateLineItemRows = 2; // rows 7 and 8

        if (numLineItems > templateLineItemRows) {
            // Need more space - shift footer down
            int extraRowsNeeded = numLineItems - templateLineItemRows;
            templateSheet.shiftRows(FORM_FOOTER_START, originalLastRow, extraRowsNeeded);
        }
        // If numLineItems <= 2, no shifting needed - maintain professional spacing

        // Step 3: Modify row 6 with actual table headers (if needed)
        // The table headers should already be properly set in the template
        // This step is reserved for future enhancement if dynamic headers are needed

        // Step 4: Copy styles from template rows 7-8 first (before overwriting)
        CellStyle[] oddRowStyles = copyCellStylesOfRow(templateSheet.getRow(ODD_STYLE_ROW));
        CellStyle[] evenRowStyles = copyCellStylesOfRow(templateSheet.getRow(EVEN_STYLE_ROW));

        // Step 5: Modify rows 7-8 with first two line items, then create additional rows
        for (int i = 0; i < numLineItems; i++) {
            Object[] lineItemData = lineItems.get(i);
            int currentRowIndex = ODD_STYLE_ROW + i; // Start from row 7

            Row lineItemRow = templateSheet.getRow(currentRowIndex);
            if (lineItemRow == null) {
                lineItemRow = templateSheet.createRow(currentRowIndex);
            }

            // Determine alternating style (0-based index: even index = odd row style)
            boolean useOddStyle = ( (i + 1) % 2 == 1 );
            CellStyle[] rowStyles = useOddStyle ? oddRowStyles : evenRowStyles;

            // Clear existing cells if this is a template row being overwritten
            if (i < templateLineItemRows) {
                for (Cell cell : lineItemRow) {
                    if (cell != null) {
                        cell.setBlank(); // Clear any existing content while preserving styles
                    }
                }
            }

            // Populate cells with data and styles
            for (int j = 0; j < lineItemData.length; j++) {
                Cell cell = lineItemRow.getCell(j);
                if (cell == null) {
                    cell = lineItemRow.createCell(j);
                }

                // Apply alternating style
                if (j < rowStyles.length && rowStyles[j] != null) {
                    cell.setCellStyle(rowStyles[j]);
                }

                // Set actual data
                setCellValue(cell, lineItemData[j]);
            }
        }

        // Step 6: Handle edge case - clear dummy data from unused template rows
        if (numLineItems < templateLineItemRows) {
            for (int i = numLineItems; i < templateLineItemRows; i++) {
                Row unusedTemplateRow = templateSheet.getRow(ODD_STYLE_ROW + i);
                if (unusedTemplateRow != null) {
                    // Clear dummy data but keep the row structure and styles for professional appearance
                    for (Cell cell : unusedTemplateRow) {
                        if (cell != null) {
                            cell.setBlank(); // Clear any content regardless of type while preserving styles
                        }
                    }
                }
            }
        }

        // Step 7: Remove any extra rows beyond the actual data (cleanup)
        int finalLastRow = templateSheet.getLastRowNum();
        int expectedLastRow = FORM_FOOTER_END + (numLineItems > templateLineItemRows ? numLineItems - templateLineItemRows : 0);

        for (int i = expectedLastRow + 1; i <= finalLastRow; i++) {
            Row row = templateSheet.getRow(i);
            if (row != null) {
                templateSheet.removeRow(row);
            }
        }

        // Step 8: Auto-size columns
        autoSizeColumns(templateSheet);
    }

    private void replacePlaceholdersInRows(Sheet sheet, int startRow, int endRow) {
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                for (Cell cell : row) {
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        String newValue = replacePlaceholders(cellValue);
                        if (!cellValue.equals(newValue)) {
                            cell.setCellValue(processStringValue(newValue, encoding));
                        }
                    }
                }
            }
        }
    }

    private CellStyle[] copyCellStylesOfRow(Row templateRow) {
        if (templateRow == null) {
            return new CellStyle[0];
        }

        int columnCount = templateRow.getLastCellNum();
        CellStyle[] cellStyles = new CellStyle[columnCount];

        for (int i = 0; i < columnCount; i++) {
            Cell cell = templateRow.getCell(i);
            if (cell != null) {
                cellStyles[i] = cell.getCellStyle();
                // Clone the style to avoid modifications affecting the template
                cellStyles[i].cloneStyleFrom(cellStyles[i]);
            }
        }
        return cellStyles;
    }

    private void autoSizeColumns(Sheet sheet) {
        if (sheet.getPhysicalNumberOfRows() > 0) {
            Row firstRow = sheet.getRow(sheet.getFirstRowNum());
            if (firstRow != null) {
                for (int i = 0; i < firstRow.getLastCellNum(); i++) {
                    sheet.autoSizeColumn(i);
                }
            }
        }
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

    private String replacePlaceholders(String text) {
        if (text == null) return text;

        String result = text;

        // Replace placeholders in format {{key}}
        for (Map.Entry<String, Object> entry : placeholderValues.entrySet()) {
            String placeholder = "{{" + entry.getKey() + "}}";
            if (result.contains(placeholder)) {
                String replacement = entry.getValue() != null ? entry.getValue().toString() : "";
                result = result.replace(placeholder, replacement);
            }
        }

        // Replace placeholders in format ${key}
        for (Map.Entry<String, Object> entry : placeholderValues.entrySet()) {
            String placeholder = "${" + entry.getKey() + "}";
            if (result.contains(placeholder)) {
                String replacement = entry.getValue() != null ? entry.getValue().toString() : "";
                result = result.replace(placeholder, replacement);
            }
        }

        return result;
    }

    private String processStringValue(String value, Charset encoding) {
        if (value == null) return null;
        // Ensure proper encoding handling
        return new String(value.getBytes(encoding), encoding);
    }

    private InputStream createTemplateStream() throws IOException {
        if (templateStream != null) {
            return templateStream;
        } else if (templatePath != null) {
            return new FileInputStream(templatePath);
        } else {
            throw new ExcelIOException("No template path or input stream provided");
        }
    }
}
