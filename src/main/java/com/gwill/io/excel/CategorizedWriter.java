package com.gwill.io.excel;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

/**
 * Specialized writer for creating categorized Excel reports with summary and detail rows.
 * This creates hierarchical reports where each category (summary row) can have multiple detail rows.
 * 
 * <h3>Example Structure</h3>
 * <pre>
 * Headers:   | Product | Q1 Sales | Q2 Sales | Total |
 * Summary:   | Laptops |   5000   |   6000   | 11000 |  ← Category row
 * Detail:    |  - Dell |   2000   |   2500   |  4500 |  ← Detail row
 * Detail:    |  - HP   |   3000   |   3500   |  6500 |  ← Detail row  
 * Summary:   | Phones  |   8000   |   9000   | 17000 |  ← Category row
 * Detail:    |  - iPhone|  5000   |   6000   | 11000 |  ← Detail row
 * Detail:    |  - Samsung| 3000   |   3000   |  6000 |  ← Detail row
 * </pre>
 * 
 * <h3>Usage Example</h3>
 * <pre>{@code
 * ExcelIO.write("sales_report.xlsx")
 *     .categorized()
 *     .sheet("Sales Report")
 *     .header("Product", "Q1", "Q2", "Total")
 *     .category("Laptops", 5000, 6000, 11000)
 *         .detail("Dell", 2000, 2500, 4500)
 *         .detail("HP", 3000, 3500, 6500)
 *     .category("Phones", 8000, 9000, 17000)
 *         .detail("iPhone", 5000, 6000, 11000)
 *         .detail("Samsung", 3000, 3000, 6000)
 *     .save();
 * }</pre>
 */
public class CategorizedWriter {
    
    private final String filePath;
    private CategorizedSheetBuilder currentSheet;
    private final List<CategorizedSheetBuilder> sheets = new ArrayList<>();
    private Charset encoding = StandardCharsets.UTF_8;

    public CategorizedWriter(String filePath) {
        this.filePath = filePath;
    }

    /**
     * Set the encoding for processing string data.
     * Default is UTF-8, which works for most international characters.
     * 
     * @param encodingName the encoding name (e.g., "UTF-8", "GBK", "ISO-8859-1")
     * @return this CategorizedWriter for method chaining
     * @throws ExcelIOException if the encoding is not supported
     */
    public CategorizedWriter encoding(String encodingName) {
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
     * @return this CategorizedWriter for method chaining
     */
    public CategorizedWriter encoding(Charset charset) {
        this.encoding = charset;
        return this;
    }

    /**
     * Add a new categorized sheet with the given name.
     * 
     * @param sheetName the name of the sheet
     * @return this CategorizedWriter for method chaining
     */
    public CategorizedWriter sheet(String sheetName) {
        currentSheet = new CategorizedSheetBuilder(sheetName, encoding);
        sheets.add(currentSheet);
        return this;
    }

    /**
     * Set the headers for the current sheet.
     * 
     * @param headers the column headers
     * @return this CategorizedWriter for method chaining
     */
    public CategorizedWriter header(String... headers) {
        ensureCurrentSheet();
        currentSheet.header(headers);
        return this;
    }

    /**
     * Start a new category (summary row) in the current sheet.
     * 
     * @param categoryValues the values for the category row
     * @return this CategorizedWriter for method chaining
     */
    public CategorizedWriter category(Object... categoryValues) {
        ensureCurrentSheet();
        currentSheet.category(categoryValues);
        return this;
    }

    /**
     * Add a detail row under the current category.
     * 
     * @param detailValues the values for the detail row
     * @return this CategorizedWriter for method chaining
     */
    public CategorizedWriter detail(Object... detailValues) {
        ensureCurrentSheet();
        currentSheet.detail(detailValues);
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
            for (CategorizedSheetBuilder sheetBuilder : sheets) {
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
     * Internal builder class for constructing categorized sheets.
     */
    static class CategorizedSheetBuilder {
        private final String sheetName;
        private final Charset encoding;
        private final List<String> headers = new ArrayList<>();
        private final List<CategoryGroup> categories = new ArrayList<>();
        private CategoryGroup currentCategory;
        private boolean headersSet = false;

        public CategorizedSheetBuilder(String sheetName, Charset encoding) {
            this.sheetName = sheetName;
            this.encoding = encoding;
        }

        public CategorizedSheetBuilder header(String... headers) {
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

        public CategorizedSheetBuilder category(Object... values) {
            // Process string values with proper encoding
            Object[] processedValues = processValues(values);
            currentCategory = new CategoryGroup(processedValues);
            categories.add(currentCategory);
            return this;
        }

        public CategorizedSheetBuilder detail(Object... values) {
            if (currentCategory == null) {
                throw new ExcelIOException("Cannot add detail row without a category. Call category() first.");
            }
            // Process string values with proper encoding
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
            // Ensure proper encoding handling
            return new String(value.getBytes(encoding), encoding);
        }

        public void buildSheet(Workbook workbook) {
            Sheet sheet = workbook.createSheet(sheetName);
            int rowNum = 0;

            // Create cell styles
            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle categoryStyle = createCategoryStyle(workbook);
            CellStyle detailStyle = createDetailStyle(workbook);

            // Write headers if present
            if (!headers.isEmpty()) {
                Row headerRow = sheet.createRow(rowNum++);
                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers.get(i));
                    cell.setCellStyle(headerStyle);
                }
            }

            // Write categories and details
            for (CategoryGroup category : categories) {
                // Write category (summary) row
                Row categoryRow = sheet.createRow(rowNum++);
                for (int i = 0; i < category.categoryValues.length; i++) {
                    Cell cell = categoryRow.createCell(i);
                    setCellValue(cell, category.categoryValues[i]);
                    cell.setCellStyle(categoryStyle);
                }

                // Write detail rows
                for (Object[] detailValues : category.detailRows) {
                    Row detailRow = sheet.createRow(rowNum++);
                    for (int i = 0; i < detailValues.length; i++) {
                        Cell cell = detailRow.createCell(i);
                        setCellValue(cell, detailValues[i]);
                        cell.setCellStyle(detailStyle);
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

        private CellStyle createHeaderStyle(Workbook workbook) {
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            font.setFontHeightInPoints((short) 12);
            style.setFont(font);
            style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            return style;
        }

        private CellStyle createCategoryStyle(Workbook workbook) {
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            font.setFontHeightInPoints((short) 11);
            style.setFont(font);
            style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            return style;
        }

        private CellStyle createDetailStyle(Workbook workbook) {
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setFontHeightInPoints((short) 10);
            style.setFont(font);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            return style;
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