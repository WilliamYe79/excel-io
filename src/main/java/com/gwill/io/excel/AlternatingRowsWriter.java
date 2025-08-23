package com.gwill.io.excel;

import com.gwill.io.excel.util.FileCopier;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * Excel writer that creates tables with alternating row styles for better readability.
 * Reads styling from a template file with 3 rows: header, odd rows, and even rows.
 *
 * <p>This writer applies different styles to alternating rows (zebra striping) which makes
 * large tables much easier to read, especially for reports with many rows of data.</p>
 *
 * <h3>Template Format</h3>
 * <p>The template Excel file should have exactly 3 rows with the desired styles:</p>
 * <ul>
 *   <li>Row 0: Header row styles</li>
 *   <li>Row 1: Odd data row styles (rows 1, 3, 5, etc.)</li>
 *   <li>Row 2: Even data row styles (rows 2, 4, 6, etc.)</li>
 * </ul>
 *
 * <h3>Usage Example</h3>
 * <pre>{@code
 * ExcelIO.alternatingRows("table_template.xlsx")
 *     .sheet("Employee Report")
 *     .header("ID", "Name", "Department", "Salary")
 *     .row("E001", "John Doe", "Engineering", 75000)
 *     .row("E002", "Jane Smith", "Marketing", 65000)
 *     .row("E003", "Mike Wilson", "Sales", 58000)
 *     .saveAs("styled_employee_report.xlsx");
 * }</pre>
 */
@RequiredArgsConstructor
public class AlternatingRowsWriter {

    public static final int TEMPLATE_HEADER_ROW_INDEX = 0;
    public static final int TEMPLATE_ODD_ROW_INDEX = 1;
    public static final int TEMPLATE_EVEN_ROW_INDEX = 2;

    private final String templatePath;
    private final InputStream templateStream;
    private AlternatingSheetBuilder currentSheet;
    private final List<AlternatingSheetBuilder> sheets = new ArrayList<>();
    private Charset encoding = StandardCharsets.UTF_8;

    public AlternatingRowsWriter(String templatePath) {
        this.templatePath = templatePath;
        this.templateStream = null;
    }

    public AlternatingRowsWriter(InputStream templateStream) {
        this.templatePath = null;
        this.templateStream = templateStream;
    }

    /**
     * Set the encoding for processing string data.
     * Default is UTF-8, which works for most international characters.
     *
     * @param encodingName the encoding name (e.g., "UTF-8", "GBK", "ISO-8859-1")
     * @return this AlternatingRowsWriter for method chaining
     * @throws ExcelIOException if the encoding is not supported
     */
    public AlternatingRowsWriter encoding(String encodingName) {
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
     * @return this AlternatingRowsWriter for method chaining
     */
    public AlternatingRowsWriter encoding(Charset charset) {
        this.encoding = charset;
        return this;
    }

    /**
     * Add a new sheet with alternating row styles.
     *
     * @param sheetName the name of the output sheet
     * @return this AlternatingRowsWriter for method chaining
     */
    public AlternatingRowsWriter sheet(String sheetName) {
        currentSheet = new AlternatingSheetBuilder(sheetName, encoding);
        sheets.add(currentSheet);
        return this;
    }

    /**
     * Add a new sheet populated from a collection of objects.
     * Headers will be automatically inferred from the first object's fields/properties.
     *
     * @param sheetName the name of the sheet
     * @param data the collection of objects to write
     * @return this AlternatingRowsWriter for method chaining
     */
    public <T> AlternatingRowsWriter sheet(String sheetName, Collection<T> data) {
        currentSheet = AlternatingSheetBuilder.fromCollection(sheetName, data, encoding);
        sheets.add(currentSheet);
        return this;
    }


    /**
     * Set the headers for the current sheet.
     *
     * @param headers the column headers
     * @return this AlternatingRowsWriter for method chaining
     */
    public AlternatingRowsWriter header(String... headers) {
        ensureCurrentSheet();
        currentSheet.header(headers);
        return this;
    }

    /**
     * Add a data row to the current sheet.
     *
     * @param values the row values
     * @return this AlternatingRowsWriter for method chaining
     */
    public AlternatingRowsWriter row(Object... values) {
        ensureCurrentSheet();
        currentSheet.row(values);
        return this;
    }

    /**
     * Add multiple data rows to the current sheet.
     *
     * @param rows the collection of row data
     * @return this AlternatingRowsWriter for method chaining
     */
    public AlternatingRowsWriter rows(Collection<Object[]> rows) {
        ensureCurrentSheet();
        rows.forEach(currentSheet::row);
        return this;
    }

    /**
     * Save the populated template to the specified file path.
     * Uses in-memory byte array operations to avoid POI style conflicts.
     *
     * @param outputPath the path where the Excel file will be saved
     * @throws ExcelIOException if there's an error processing the template or saving the file
     */
    public void saveAs(String outputPath) {
        if (sheets.isEmpty()) {
            throw new ExcelIOException("No sheets defined. Call sheet() method first.");
        }

        try {
            byte[] resultBytes = processTemplateInMemory();
            FileCopier.writeByteArrayToFile(resultBytes, outputPath);
        } catch (IOException e) {
            throw new ExcelIOException("Failed to process template and save Excel file: " + outputPath, e);
        }
    }

    /**
     * Write the populated template to the specified output stream.
     * This method is ideal for web applications and HTTP responses.
     *
     * @param outputStream the output stream to write to
     * @throws ExcelIOException if there's an error processing the template or writing to the stream
     */
    public void writeTo(OutputStream outputStream) {
        if (sheets.isEmpty()) {
            throw new ExcelIOException("No sheets defined. Call sheet() method first.");
        }
        if (outputStream == null) {
            throw new IllegalArgumentException("Output stream cannot be null");
        }

        try {
            byte[] resultBytes = processTemplateInMemory();
            outputStream.write(resultBytes);
            outputStream.flush();
        } catch (IOException e) {
            throw new ExcelIOException("Failed to process template and write to output stream", e);
        }
    }

    /**
     * Return the populated template as a byte array.
     * Uses in-memory operations to avoid POI style conflicts.
     *
     * @return the populated Excel file as a byte array
     * @throws ExcelIOException if there's an error processing the template
     */
    public byte[] toByteArray() {
        if (sheets.isEmpty()) {
            throw new ExcelIOException("No sheets defined. Call sheet() method first.");
        }

        try {
            return processTemplateInMemory();
        } catch (IOException e) {
            throw new ExcelIOException("Failed to process template to byte array", e);
        }
    }

    private void ensureCurrentSheet() {
        if (currentSheet == null) {
            sheet("Sheet1"); // Default sheet name
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

                // 1. Get and store the styles from template
                Sheet templateSheet = workbook.getSheetAt(0);
                StylesHolder stylesHolder = readStylesFromTemplate(templateSheet, workbook);

                // 2. Check if we have any sheets to process
                if (sheets.isEmpty()) {
                    throw new ExcelIOException("No sheets to process");
                }

                // 3. Get the first sheet and populate the template sheet with its data
                AlternatingSheetBuilder firstSheetBuilder = sheets.get(0);

                // Rename the template sheet if needed
                if (!firstSheetBuilder.sheetName.equals(templateSheet.getSheetName())) {
                    workbook.setSheetName(0, firstSheetBuilder.sheetName);
                }

                // Populate the template sheet with first sheet's data
                firstSheetBuilder.populateTemplateSheet(templateSheet);

                // 4. If we have more sheets, create new ones with stored styles
                for (int i = 1; i < sheets.size(); i++) {
                    AlternatingSheetBuilder sheetBuilder = sheets.get(i);
                    sheetBuilder.buildSheet(workbook, stylesHolder);
                }

                // Return as byte array
                try (ByteArrayOutputStream byteOut = new ByteArrayOutputStream()) {
                    workbook.write(byteOut);
                    return byteOut.toByteArray();
                }
            }
        }
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

    private StylesHolder readStylesFromTemplate(Sheet templateSheet, Workbook templateWorkbook) {
        StylesHolder styles = new StylesHolder();

        // Read header row styles (row 0)
        Row headerRow = templateSheet.getRow(TEMPLATE_HEADER_ROW_INDEX);
        if (headerRow != null) {
            styles.headerStyles = readRowStyles(headerRow, templateWorkbook);
        }

        // Read odd row styles (row 1)
        Row oddRow = templateSheet.getRow(TEMPLATE_ODD_ROW_INDEX);
        if (oddRow != null) {
            styles.oddRowStyles = readRowStyles(oddRow, templateWorkbook);
        }

        // Read even row styles (row 2)
        Row evenRow = templateSheet.getRow(TEMPLATE_EVEN_ROW_INDEX);
        if (evenRow != null) {
            styles.evenRowStyles = readRowStyles(evenRow, templateWorkbook);
        }

        return styles;
    }

    private List<CellStyle> readRowStyles(Row templateRow, Workbook outputWorkbook) {
        List<CellStyle> rowStyles = new ArrayList<>();

        for (Cell templateCell : templateRow) {
            if (templateCell != null) {
                // Clone the cell style to the output workbook
                CellStyle newStyle = outputWorkbook.createCellStyle();
                newStyle.cloneStyleFrom(templateCell.getCellStyle());
                rowStyles.add(newStyle);
            } else {
                rowStyles.add(null);
            }
        }

        return rowStyles;
    }

    /**
     * Holder for the three types of styles read from template.
     */
    static class StylesHolder {
        List<CellStyle> headerStyles = new ArrayList<>();
        List<CellStyle> oddRowStyles = new ArrayList<>();
        List<CellStyle> evenRowStyles = new ArrayList<>();
    }

    /**
     * Internal builder class for constructing sheets with alternating row styles.
     */
    static class AlternatingSheetBuilder {
        private final String sheetName;
        private final Charset encoding;
        private final List<String> headers = new ArrayList<>();
        private final List<Object[]> rows = new ArrayList<>();
        private boolean headersSet = false;

        public AlternatingSheetBuilder(String sheetName, Charset encoding) {
            this.sheetName = sheetName;
            this.encoding = encoding;
        }

        public static <T> AlternatingSheetBuilder fromCollection(String sheetName, Collection<T> data, Charset encoding) {
            AlternatingSheetBuilder builder = new AlternatingSheetBuilder(sheetName, encoding);
            if (!data.isEmpty()) {
                T first = data.iterator().next();
                if (first instanceof Map) {
                    return fromMaps(sheetName, (Collection<Map<String, Object>>) data, encoding);
                }
                // For other objects, just convert to string representation
                builder.header("Data");
                for (T item : data) {
                    builder.row(processStringValue(item.toString(), encoding));
                }
            }
            return builder;
        }

        public static AlternatingSheetBuilder fromMaps(String sheetName, Collection<Map<String, Object>> data, Charset encoding) {
            AlternatingSheetBuilder builder = new AlternatingSheetBuilder(sheetName, encoding);
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
            return new String(value.getBytes(encoding), encoding);
        }

        public AlternatingSheetBuilder header(String... headers) {
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

        public AlternatingSheetBuilder row(Object... values) {
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

        /**
         * Populate an existing template sheet with data, preserving styles from template rows.
         * This follows the pattern from your existing implementation.
         */
        public void populateTemplateSheet(Sheet templateSheet) {
            // Verify template has the required rows
            if (templateSheet.getLastRowNum() < TEMPLATE_EVEN_ROW_INDEX) {
                throw new ExcelIOException("Template sheet must have at least 3 rows (header, odd style, even style)");
            }

            int currentRowNum = 0;

            // Write headers if present (modify row 0)
            if (!headers.isEmpty()) {
                Row headerRow = templateSheet.getRow(TEMPLATE_HEADER_ROW_INDEX);
                if (headerRow == null) {
                    headerRow = templateSheet.createRow(TEMPLATE_HEADER_ROW_INDEX);
                }

                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = headerRow.getCell(i);
                    if (cell == null) {
                        cell = headerRow.createCell(i);
                        // Copy style from first column if this is a new cell
                        if (i == 0 && headerRow.getCell(0) != null) {
                            // Style is already there from template
                        }
                    }
                    cell.setCellValue(headers.get(i));
                }
                currentRowNum = 1;
            }

            // Copy style patterns from template rows
            CellStyle[] oddRowStyles = copyCellStylesOfRow(templateSheet.getRow(TEMPLATE_ODD_ROW_INDEX));
            CellStyle[] evenRowStyles = copyCellStylesOfRow(templateSheet.getRow(TEMPLATE_EVEN_ROW_INDEX));

            // Write data rows starting after the template rows
            int dataRowNum = 1;

            for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
                Object[] rowData = rows.get(rowIndex);
                Row dataRow = templateSheet.getRow(dataRowNum);
                if (dataRow == null) {
                    dataRow = templateSheet.createRow(dataRowNum);
                }

                // Determine if this is an odd or even data row
                boolean isEvenRow = ( (rowIndex + 1) % 2 == 0 );
                CellStyle[] rowStyles = isEvenRow ? evenRowStyles : oddRowStyles;

                for (int i = 0; i < rowData.length; i++) {
                    Cell cell = dataRow.getCell(i);
                    if (cell == null) {
                        cell = dataRow.createCell(i);
                    }

                    // Apply alternating style
                    if (i < rowStyles.length && rowStyles[i] != null) {
                        cell.setCellStyle(rowStyles[i]);
                    }

                    setCellValue(cell, rowData[i]);
                }
                dataRowNum++;
            }

            // Remove any extra rows beyond our data
            int lastRowNum = templateSheet.getLastRowNum();
            for (int i = dataRowNum; i <= lastRowNum; i++) {
                Row row = templateSheet.getRow(i);
                if (row != null) {
                    templateSheet.removeRow(row);
                }
            }

            // Auto-size columns
            int maxColumns = Math.max(headers.size(), getMaxColumnCount());
            for (int i = 0; i < maxColumns; i++) {
                templateSheet.autoSizeColumn(i);
            }
        }

        /**
         * Copy cell styles from a template row, following your existing pattern.
         */
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

        public void buildSheet(Workbook workbook, StylesHolder stylesHolder) {
            Sheet sheet = workbook.createSheet(sheetName);
            int rowNum = 0;

            // Write headers if present
            if (!headers.isEmpty()) {
                Row headerRow = sheet.createRow(rowNum++);
                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers.get(i));

                    // Apply header style from template if available
                    if (i < stylesHolder.headerStyles.size() && stylesHolder.headerStyles.get(i) != null) {
                        cell.setCellStyle(stylesHolder.headerStyles.get(i));
                    }
                }
            }

            // Write data rows with alternating styles
            for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
                Object[] rowData = rows.get(rowIndex);
                Row dataRow = sheet.createRow(rowNum++);

                // Determine if this is an odd or even data row (0-based, so first data row is index 0 = even)
                boolean isEvenRow = (rowIndex % 2 == 0);
                List<CellStyle> rowStyles = isEvenRow ? stylesHolder.evenRowStyles : stylesHolder.oddRowStyles;

                for (int i = 0; i < rowData.length; i++) {
                    Cell cell = dataRow.createCell(i);
                    setCellValue(cell, rowData[i]);

                    // Apply alternating style from template if available
                    if (i < rowStyles.size() && rowStyles.get(i) != null) {
                        cell.setCellStyle(rowStyles.get(i));
                    }
                }
            }

            // Auto-size columns
            int maxColumns = Math.max(headers.size(), getMaxColumnCount());
            for (int i = 0; i < maxColumns; i++) {
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
