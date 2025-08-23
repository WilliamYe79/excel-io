package com.gwill.io.excel;

import com.gwill.io.excel.util.FileCopier;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

/**
 * Template-based categorized Excel writer that reads cell styles from a template file.
 * This creates hierarchical reports where styles are preserved from the template.
 *
 * <p>The template file should have exactly 3 rows with the desired styles:</p>
 * <ul>
 *   <li>Row 0: Header row styles</li>
 *   <li>Row 1: Category/summary row styles</li>
 *   <li>Row 2: Detail row styles</li>
 * </ul>
 *
 * <h3>Usage Example</h3>
 * <pre>{@code
 * ExcelIO.categorizedFromTemplate("template.xlsx")
 *     .sheet("Sales Report")
 *     .header("Product", "Q1", "Q2", "Total")
 *     .category("Laptops", 5000, 6000, 11000)
 *         .detail("Dell", 2000, 2500, 4500)
 *         .detail("HP", 3000, 3500, 6500)
 *     .saveAs("output.xlsx");
 * }</pre>
 */
@RequiredArgsConstructor
public class CategorizedTemplateWriter {

    public static final int TEMPLATE_HEADINGS_ROW_INDEX = 0;
    public static final int TEMPLATE_SUMMARY_ROW_INDEX = 1;
    public static final int TEMPLATE_DETAILED_ROW_INDEX = 2;

    private final String templatePath;
    private final InputStream templateStream;
    private CategorizedTemplateSheetBuilder currentSheet;
    private final List<CategorizedTemplateSheetBuilder> sheets = new ArrayList<>();
    private Charset encoding = StandardCharsets.UTF_8;

    public CategorizedTemplateWriter(String templatePath) {
        this.templatePath = templatePath;
        this.templateStream = null;
    }

    public CategorizedTemplateWriter(InputStream templateStream) {
        this.templatePath = null;
        this.templateStream = templateStream;
    }

    /**
     * Set the encoding for processing string data.
     * Default is UTF-8, which works for most international characters.
     *
     * @param encodingName the encoding name (e.g., "UTF-8", "GBK", "ISO-8859-1")
     * @return this CategorizedTemplateWriter for method chaining
     * @throws ExcelIOException if the encoding is not supported
     */
    public CategorizedTemplateWriter encoding(String encodingName) {
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
     * @return this CategorizedTemplateWriter for method chaining
     */
    public CategorizedTemplateWriter encoding(Charset charset) {
        this.encoding = charset;
        return this;
    }

    /**
     * Add a new categorized sheet with the given name.
     * This will use the first sheet of the template for styling.
     *
     * @param sheetName the name of the output sheet
     * @return this CategorizedTemplateWriter for method chaining
     */
    public CategorizedTemplateWriter sheet(String sheetName) {
        currentSheet = new CategorizedTemplateSheetBuilder(sheetName, encoding);
        sheets.add(currentSheet);
        return this;
    }

    /**
     * Set the headers for the current sheet.
     *
     * @param headers the column headers
     * @return this CategorizedTemplateWriter for method chaining
     */
    public CategorizedTemplateWriter header(String... headers) {
        ensureCurrentSheet();
        currentSheet.header(headers);
        return this;
    }

    /**
     * Start a new category (summary row) in the current sheet.
     *
     * @param categoryValues the values for the category row
     * @return this CategorizedTemplateWriter for method chaining
     */
    public CategorizedTemplateWriter category(Object... categoryValues) {
        ensureCurrentSheet();
        currentSheet.category(categoryValues);
        return this;
    }

    /**
     * Add a detail row under the current category.
     *
     * @param detailValues the values for the detail row
     * @return this CategorizedTemplateWriter for method chaining
     */
    public CategorizedTemplateWriter detail(Object... detailValues) {
        ensureCurrentSheet();
        currentSheet.detail(detailValues);
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
                CategorizedTemplateSheetBuilder firstSheetBuilder = sheets.get(0);

                // Rename the template sheet if needed
                if (!firstSheetBuilder.sheetName.equals(templateSheet.getSheetName())) {
                    workbook.setSheetName(0, firstSheetBuilder.sheetName);
                }

                // Populate the template sheet with first sheet's data
                firstSheetBuilder.populateTemplateSheet(templateSheet);

                // 4. If we have more sheets, create new ones with stored styles
                for (int i = 1; i < sheets.size(); i++) {
                    CategorizedTemplateSheetBuilder sheetBuilder = sheets.get(i);
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

    private void ensureCurrentSheet() {
        if (currentSheet == null) {
            sheet("Sheet1"); // Default sheet name
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
        Row headerRow = templateSheet.getRow(TEMPLATE_HEADINGS_ROW_INDEX);
        if (headerRow != null) {
            styles.headerStyles = readRowStyles(headerRow, templateWorkbook);
        }

        // Read summary/category row styles (row 1)
        Row summaryRow = templateSheet.getRow(TEMPLATE_SUMMARY_ROW_INDEX);
        if (summaryRow != null) {
            styles.categoryStyles = readRowStyles(summaryRow, templateWorkbook);
        }

        // Read detail row styles (row 2)
        Row detailRow = templateSheet.getRow(TEMPLATE_DETAILED_ROW_INDEX);
        if (detailRow != null) {
            styles.detailStyles = readRowStyles(detailRow, templateWorkbook);
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
        List<CellStyle> categoryStyles = new ArrayList<>();
        List<CellStyle> detailStyles = new ArrayList<>();
    }

    /**
     * Internal builder class for constructing categorized sheets with template styles.
     */
    static class CategorizedTemplateSheetBuilder {
        private final String sheetName;
        private final Charset encoding;
        private final List<String> headers = new ArrayList<>();
        private final List<CategoryGroup> categories = new ArrayList<>();
        private CategoryGroup currentCategory;
        private boolean headersSet = false;

        public CategorizedTemplateSheetBuilder(String sheetName, Charset encoding) {
            this.sheetName = sheetName;
            this.encoding = encoding;
        }

        public CategorizedTemplateSheetBuilder header(String... headers) {
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

        public CategorizedTemplateSheetBuilder category(Object... values) {
            Object[] processedValues = processValues(values);
            currentCategory = new CategoryGroup(processedValues);
            categories.add(currentCategory);
            return this;
        }

        public CategorizedTemplateSheetBuilder detail(Object... values) {
            if (currentCategory == null) {
                throw new ExcelIOException("Cannot add detail row without a category. Call category() first.");
            }
            Object[] processedValues = processValues(values);
            currentCategory.addDetail(processedValues);
            return this;
        }

        private Object[] processValues(Object[] values) {
            Object[] processedValues = new Object[values.length];
            for (int i = 0; i < values.length; i++) {
                if (values[i] instanceof String) {
                    processedValues[i] = processStringValue((String) values[i], encoding);
                } else {
                    processedValues[i] = values[i];
                }
            }
            return processedValues;
        }

        private static String processStringValue(String value, Charset encoding) {
            if (value == null) return null;
            return new String(value.getBytes(encoding), encoding);
        }

        /**
         * Populate an existing template sheet with categorized data, preserving styles from template rows.
         * This follows the pattern from your existing implementation - starts from row 1 and overwrites template data.
         */
        public void populateTemplateSheet(Sheet templateSheet) {
            // Verify template has the required rows
            if (templateSheet.getLastRowNum() < TEMPLATE_DETAILED_ROW_INDEX) {
                throw new ExcelIOException("Template sheet must have at least 3 rows (header, category, detail)");
            }

            int lastRowNumber = templateSheet.getLastRowNum();

            // Write headers if present (modify row 0)
            if (!headers.isEmpty()) {
                Row headerRow = templateSheet.getRow(TEMPLATE_HEADINGS_ROW_INDEX);
                if (headerRow == null) {
                    headerRow = templateSheet.createRow(TEMPLATE_HEADINGS_ROW_INDEX);
                }

                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = headerRow.getCell(i);
                    if (cell == null) {
                        cell = headerRow.createCell(i);
                    }
                    cell.setCellValue(headers.get(i));
                }
            }

            // Start writing data from row 1 (this will overwrite template rows naturally)
            int rowIndex = 1;

            // Write categories and details
            for (CategoryGroup category : categories) {
                // Write category (summary) row - this overwrites template summary row
                Row categoryRow = templateSheet.getRow(rowIndex);
                if (categoryRow == null) {
                    categoryRow = templateSheet.createRow(rowIndex);
                }
                writeCategoryRowWithCellStyles(templateSheet, categoryRow, category.categoryValues);
                rowIndex++;

                // Write detail rows - these overwrite template detail row and beyond
                for (Object[] detailValues : category.detailRows) {
                    Row detailRow = templateSheet.getRow(rowIndex);
                    if (detailRow == null) {
                        detailRow = templateSheet.createRow(rowIndex);
                    }
                    writeDetailRowWithCellStyles(templateSheet, detailRow, detailValues);
                    rowIndex++;
                }
            }

            // Remove extra rows if any (following your pattern exactly)
            for (int i = rowIndex; i <= lastRowNumber; i++) {
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

        private void writeCategoryRowWithCellStyles(Sheet sheet, Row row, Object[] rowData) {
            CellStyle[] categoryStyles = copyCellStylesOfRow(sheet.getRow(TEMPLATE_SUMMARY_ROW_INDEX));

            for (int i = 0; i < rowData.length; i++) {
                Cell cell = row.getCell(i);
                if (cell == null) {
                    cell = row.createCell(i);
                }

                // Apply category style
                if (i < categoryStyles.length && categoryStyles[i] != null) {
                    cell.setCellStyle(categoryStyles[i]);
                }

                setCellValue(cell, rowData[i]);
            }
        }

        private void writeDetailRowWithCellStyles(Sheet sheet, Row row, Object[] rowData) {
            CellStyle[] detailStyles = copyCellStylesOfRow(sheet.getRow(TEMPLATE_DETAILED_ROW_INDEX));

            for (int i = 0; i < rowData.length; i++) {
                Cell cell = row.getCell(i);
                if (cell == null) {
                    cell = row.createCell(i);
                }

                // Apply detail style
                if (i < detailStyles.length && detailStyles[i] != null) {
                    cell.setCellStyle(detailStyles[i]);
                }

                setCellValue(cell, rowData[i]);
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

            // Write categories and details
            for (CategoryGroup category : categories) {
                // Write category (summary) row
                Row categoryRow = sheet.createRow(rowNum++);
                for (int i = 0; i < category.categoryValues.length; i++) {
                    Cell cell = categoryRow.createCell(i);
                    setCellValue(cell, category.categoryValues[i]);

                    // Apply category style from template if available
                    if (i < stylesHolder.categoryStyles.size() && stylesHolder.categoryStyles.get(i) != null) {
                        cell.setCellStyle(stylesHolder.categoryStyles.get(i));
                    }
                }

                // Write detail rows
                for (Object[] detailValues : category.detailRows) {
                    Row detailRow = sheet.createRow(rowNum++);
                    for (int i = 0; i < detailValues.length; i++) {
                        Cell cell = detailRow.createCell(i);
                        setCellValue(cell, detailValues[i]);

                        // Apply detail style from template if available
                        if (i < stylesHolder.detailStyles.size() && stylesHolder.detailStyles.get(i) != null) {
                            cell.setCellStyle(stylesHolder.detailStyles.get(i));
                        }
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
            return categories.stream()
                .mapToInt(cat -> Math.max(
                    cat.categoryValues.length,
                    cat.detailRows.stream()
                        .mapToInt(row -> row.length)
                        .max()
                        .orElse(0)
                ))
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

        /**
         * Represents a category with its associated detail rows.
         */
        static class CategoryGroup {
            final Object[] categoryValues;
            final List<Object[]> detailRows = new ArrayList<>();

            public CategoryGroup(Object[] categoryValues) {
                this.categoryValues = categoryValues;
            }

            public void addDetail(Object[] detailValues) {
                detailRows.add(detailValues);
            }
        }
    }
}
