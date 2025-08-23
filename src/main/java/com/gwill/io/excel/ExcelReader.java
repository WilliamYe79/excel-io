package com.gwill.io.excel;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.*;

/**
 * Fluent builder for reading XLSX Excel files with both basic and type-safe reading capabilities.
 * Automatically detects whether to use typed reading based on metadata availability.
 * 
 * <p><strong>Note:</strong> This library supports .xlsx files only. For .xls files, 
 * please open them in Excel and save as .xlsx format first.</p>
 * 
 * <h3>Basic Reading</h3>
 * <pre>{@code
 * List<Map<String, Object>> data = ExcelIO.read("input.xlsx")
 *     .sheet("Data")
 *     .asMaps();
 * }</pre>
 * 
 * <h3>Type-Safe Reading with Metadata</h3>
 * <pre>{@code
 * List<Map<String, Object>> typedData = ExcelIO.read("data.xlsx")
 *     .withMetadata("metadata.xlsx")
 *     .sheet("Data")
 *     .asMaps();
 * }</pre>
 * 
 * <h3>Inline Type Specification</h3>
 * <pre>{@code
 * List<Map<String, Object>> typedData = ExcelIO.read("data.xlsx")
 *     .withTypes("name:java.lang.String", "age:java.lang.Integer")
 *     .sheet("Data")
 *     .asMaps();
 * }</pre>
 */
@RequiredArgsConstructor
public class ExcelReader {
    
    private final String filePath;
    private final InputStream inputStream;
    private String selectedSheet;
    private boolean hasHeaders = true;
    private ColumnMetadata columnMetadata;

    public ExcelReader(String filePath) {
        this.filePath = filePath;
        this.inputStream = null;
    }

    public ExcelReader(InputStream inputStream) {
        this.filePath = null;
        this.inputStream = inputStream;
    }

    /**
     * Specify a metadata Excel file that contains column names and types.
     * 
     * @param metadataFilePath path to the metadata Excel file
     * @return this ExcelReader for method chaining
     */
    public ExcelReader withMetadata(String metadataFilePath) {
        this.columnMetadata = readMetadataFromFile(metadataFilePath);
        return this;
    }

    /**
     * Specify a metadata InputStream that contains column names and types.
     * 
     * @param metadataStream InputStream containing metadata
     * @return this ExcelReader for method chaining
     */
    public ExcelReader withMetadata(InputStream metadataStream) {
        this.columnMetadata = readMetadataFromStream(metadataStream);
        return this;
    }

    /**
     * Specify column types inline using "columnName:java.lang.ClassName" format.
     * 
     * @param columnDefinitions array of "name:java.lang.ClassName" strings
     * @return this ExcelReader for method chaining
     */
    public ExcelReader withTypes(String... columnDefinitions) {
        this.columnMetadata = parseInlineMetadata(columnDefinitions);
        return this;
    }

    /**
     * Select a specific sheet to read from.
     * 
     * @param sheetName the name of the sheet to read
     * @return this ExcelReader for method chaining
     */
    public ExcelReader sheet(String sheetName) {
        this.selectedSheet = sheetName;
        return this;
    }

    /**
     * Select a specific sheet by index to read from.
     * 
     * @param sheetIndex the zero-based index of the sheet to read
     * @return this ExcelReader for method chaining
     */
    public ExcelReader sheet(int sheetIndex) {
        this.selectedSheet = String.valueOf(sheetIndex);
        return this;
    }

    /**
     * Specify whether the first row contains headers.
     * Default is true.
     * 
     * @param hasHeaders true if first row contains headers, false otherwise
     * @return this ExcelReader for method chaining
     */
    public ExcelReader hasHeaders(boolean hasHeaders) {
        this.hasHeaders = hasHeaders;
        return this;
    }

    /**
     * Read the data as a list of maps. Uses type conversion if metadata is available.
     * 
     * @return list of maps representing the data
     * @throws ExcelIOException if there's an error reading the file
     */
    public List<Map<String, Object>> asMaps() {
        try (Workbook workbook = createWorkbook()) {
            Sheet sheet = getSheet(workbook);
            
            if (columnMetadata != null) {
                return readAsTypedMaps(sheet);
            } else {
                return readAsBasicMaps(sheet);
            }
        } catch (IOException e) {
            throw new ExcelIOException("Failed to read Excel file", e);
        }
    }

    /**
     * Read the data as a list of object arrays.
     * 
     * @return list of object arrays representing the data
     * @throws ExcelIOException if there's an error reading the file
     */
    public List<Object[]> asRows() {
        try (Workbook workbook = createWorkbook()) {
            Sheet sheet = getSheet(workbook);
            
            if (columnMetadata != null) {
                return readAsTypedRows(sheet);
            } else {
                return readAsBasicRows(sheet);
            }
        } catch (IOException e) {
            throw new ExcelIOException("Failed to read Excel file", e);
        }
    }

    /**
     * Read all sheets as maps, where the key is the sheet name and 
     * the value is a list of maps representing the data.
     * 
     * @return map of sheet name to data
     * @throws ExcelIOException if there's an error reading the file
     */
    public Map<String, List<Map<String, Object>>> allSheetsAsMaps() {
        try (Workbook workbook = createWorkbook()) {
            Map<String, List<Map<String, Object>>> result = new LinkedHashMap<>();
            
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();
                
                if (columnMetadata != null) {
                    result.put(sheetName, readAsTypedMaps(sheet));
                } else {
                    result.put(sheetName, readAsBasicMaps(sheet));
                }
            }
            
            return result;
        } catch (IOException e) {
            throw new ExcelIOException("Failed to read Excel file", e);
        }
    }

    /**
     * Get the names of all sheets in the workbook.
     * 
     * @return list of sheet names
     * @throws ExcelIOException if there's an error reading the file
     */
    public List<String> getSheetNames() {
        try (Workbook workbook = createWorkbook()) {
            List<String> sheetNames = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames.add(workbook.getSheetAt(i).getSheetName());
            }
            return sheetNames;
        } catch (IOException e) {
            throw new ExcelIOException("Failed to read Excel file", e);
        }
    }

    // Private methods for metadata handling
    private ColumnMetadata readMetadataFromFile(String metadataFilePath) {
        try (FileInputStream fis = new FileInputStream(metadataFilePath)) {
            return readMetadataFromStream(fis);
        } catch (IOException e) {
            throw new ExcelIOException("Failed to read metadata file: " + metadataFilePath, e);
        }
    }

    private ColumnMetadata readMetadataFromStream(InputStream metadataStream) {
        try (Workbook workbook = new XSSFWorkbook(metadataStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            
            // Read column names from first row
            Row nameRow = sheet.getRow(0);
            if (nameRow == null) {
                throw new ExcelIOException("Metadata file must have at least 2 rows: names and types");
            }
            
            List<String> columnNames = new ArrayList<>();
            for (Cell cell : nameRow) {
                columnNames.add(cell != null ? cell.toString() : "");
            }
            
            // Read column types from second row
            Row typeRow = sheet.getRow(1);
            if (typeRow == null) {
                throw new ExcelIOException("Metadata file must have at least 2 rows: names and types");
            }
            
            List<ExcelJavaDataType> columnTypes = new ArrayList<>();
            for (int i = 0; i < columnNames.size(); i++) {
                Cell cell = typeRow.getCell(i);
                String typeStr = cell != null ? cell.toString().trim() : "java.lang.String";
                try {
                    ExcelJavaDataType dataType = ExcelJavaDataType.fromClassName(typeStr);
                    columnTypes.add(dataType);
                } catch (IllegalArgumentException e) {
                    throw new ExcelIOException("Unsupported data type: " + typeStr + " at column " + i + ". " + e.getMessage(), e);
                }
            }
            
            return new ColumnMetadata(columnNames, columnTypes);
            
        } catch (IOException e) {
            throw new ExcelIOException("Failed to read metadata stream", e);
        }
    }

    private ColumnMetadata parseInlineMetadata(String[] columnDefinitions) {
        List<String> columnNames = new ArrayList<>();
        List<ExcelJavaDataType> columnTypes = new ArrayList<>();
        
        for (String definition : columnDefinitions) {
            String[] parts = definition.split(":");
            if (parts.length != 2) {
                throw new ExcelIOException("Invalid column definition: " + definition + ". Use 'name:java.lang.ClassName' format.");
            }
            
            columnNames.add(parts[0].trim());
            try {
                ExcelJavaDataType dataType = ExcelJavaDataType.fromClassName(parts[1].trim());
                columnTypes.add(dataType);
            } catch (IllegalArgumentException e) {
                throw new ExcelIOException("Unsupported data type: " + parts[1] + " in definition: " + definition + ". " + e.getMessage(), e);
            }
        }
        
        return new ColumnMetadata(columnNames, columnTypes);
    }

    // Private methods for workbook operations
    private Workbook createWorkbook() throws IOException {
        if (inputStream != null) {
            return new XSSFWorkbook(inputStream);
        } else if (filePath != null) {
            return new XSSFWorkbook(new FileInputStream(filePath));
        } else {
            throw new ExcelIOException("No file path or input stream provided");
        }
    }

    private Sheet getSheet(Workbook workbook) {
        if (selectedSheet == null) {
            return workbook.getSheetAt(0);
        }

        Sheet sheet = workbook.getSheet(selectedSheet);
        if (sheet != null) {
            return sheet;
        }

        try {
            int index = Integer.parseInt(selectedSheet);
            return workbook.getSheetAt(index);
        } catch (IllegalArgumentException e) {
            throw new ExcelIOException("Sheet not found: " + selectedSheet);
        }
    }

    // Basic reading methods (from original ExcelReader)
    private List<Map<String, Object>> readAsBasicMaps(Sheet sheet) {
        List<Map<String, Object>> result = new ArrayList<>();
        List<String> headers = new ArrayList<>();
        
        Iterator<Row> rowIterator = sheet.iterator();
        boolean isFirstRow = true;

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            
            if (isFirstRow && hasHeaders) {
                headers = readRowAsStrings(row);
                isFirstRow = false;
                continue;
            } else if (isFirstRow && !hasHeaders) {
                for (int i = 0; i < row.getLastCellNum(); i++) {
                    headers.add("Column" + (i + 1));
                }
                isFirstRow = false;
            }

            Map<String, Object> rowMap = new LinkedHashMap<>();
            for (int i = 0; i < Math.max(headers.size(), row.getLastCellNum()); i++) {
                String header = i < headers.size() ? headers.get(i) : "Column" + (i + 1);
                Object value = i < row.getLastCellNum() ? getCellValue(row.getCell(i)) : null;
                rowMap.put(header, value);
            }
            result.add(rowMap);
        }

        return result;
    }

    private List<Object[]> readAsBasicRows(Sheet sheet) {
        List<Object[]> result = new ArrayList<>();
        
        for (Row row : sheet) {
            List<Object> rowData = new ArrayList<>();
            for (int i = 0; i < row.getLastCellNum(); i++) {
                rowData.add(getCellValue(row.getCell(i)));
            }
            result.add(rowData.toArray());
        }

        return result;
    }

    // Typed reading methods (from TypedExcelReader)
    private List<Map<String, Object>> readAsTypedMaps(Sheet sheet) {
        List<Map<String, Object>> result = new ArrayList<>();
        
        Iterator<Row> rowIterator = sheet.iterator();
        boolean isFirstRow = true;

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            
            if (isFirstRow && hasHeaders) {
                isFirstRow = false;
                continue;
            } else if (isFirstRow) {
                isFirstRow = false;
            }

            Map<String, Object> rowMap = new LinkedHashMap<>();
            for (int i = 0; i < columnMetadata.getColumnNames().size(); i++) {
                String columnName = columnMetadata.getColumnNames().get(i);
                ExcelJavaDataType dataType = columnMetadata.getColumnTypes().get(i);
                
                Cell cell = i < row.getLastCellNum() ? row.getCell(i) : null;
                Object typedValue = convertCellValue(cell, dataType);
                rowMap.put(columnName, typedValue);
            }
            result.add(rowMap);
        }

        return result;
    }

    private List<Object[]> readAsTypedRows(Sheet sheet) {
        List<Object[]> result = new ArrayList<>();
        
        Iterator<Row> rowIterator = sheet.iterator();
        boolean isFirstRow = true;

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            
            if (isFirstRow && hasHeaders) {
                isFirstRow = false;
                continue;
            } else if (isFirstRow) {
                isFirstRow = false;
            }

            Object[] typedRow = new Object[columnMetadata.getColumnTypes().size()];
            for (int i = 0; i < columnMetadata.getColumnTypes().size(); i++) {
                ExcelJavaDataType dataType = columnMetadata.getColumnTypes().get(i);
                Cell cell = i < row.getLastCellNum() ? row.getCell(i) : null;
                typedRow[i] = convertCellValue(cell, dataType);
            }
            result.add(typedRow);
        }

        return result;
    }

    // Utility methods
    private List<String> readRowAsStrings(Row row) {
        List<String> result = new ArrayList<>();
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            result.add(cell != null ? cell.toString() : "");
        }
        return result;
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    double numValue = cell.getNumericCellValue();
                    if (numValue == Math.floor(numValue)) {
                        return (long) numValue;
                    } else {
                        return numValue;
                    }
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                try {
                    return cell.getNumericCellValue();
                } catch (Exception e) {
                    return cell.getStringCellValue();
                }
            case BLANK:
            case _NONE:
            default:
                return null;
        }
    }

    private Object convertCellValue(Cell cell, ExcelJavaDataType targetType) {
        if (cell == null) {
            return null;
        }

        try {
            switch (targetType) {
                case STRING:
                    return cell.toString();
                    
                case BOOLEAN:
                    if (cell.getCellType() == CellType.BOOLEAN) {
                        return cell.getBooleanCellValue();
                    }
                    return Boolean.parseBoolean(cell.toString());
                    
                case INTEGER:
                    if (cell.getCellType() == CellType.NUMERIC) {
                        return (int) cell.getNumericCellValue();
                    }
                    return Integer.parseInt(cell.toString());
                    
                case LONG:
                    if (cell.getCellType() == CellType.NUMERIC) {
                        return (long) cell.getNumericCellValue();
                    }
                    return Long.parseLong(cell.toString());
                    
                case SHORT:
                    if (cell.getCellType() == CellType.NUMERIC) {
                        return (short) cell.getNumericCellValue();
                    }
                    return Short.parseShort(cell.toString());
                    
                case BYTE:
                    if (cell.getCellType() == CellType.NUMERIC) {
                        return (byte) cell.getNumericCellValue();
                    }
                    return Byte.parseByte(cell.toString());
                    
                case FLOAT:
                    if (cell.getCellType() == CellType.NUMERIC) {
                        return (float) cell.getNumericCellValue();
                    }
                    return Float.parseFloat(cell.toString());
                    
                case DOUBLE:
                    if (cell.getCellType() == CellType.NUMERIC) {
                        return cell.getNumericCellValue();
                    }
                    return Double.parseDouble(cell.toString());
                    
                case BIG_DECIMAL:
                    if (cell.getCellType() == CellType.NUMERIC) {
                        return BigDecimal.valueOf(cell.getNumericCellValue());
                    }
                    return new BigDecimal(cell.toString());
                    
                case LOCAL_DATE:
                    if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                        return cell.getLocalDateTimeCellValue().toLocalDate();
                    }
                    return LocalDate.parse(cell.toString());
                    
                case LOCAL_TIME:
                    if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                        return cell.getLocalDateTimeCellValue().toLocalTime();
                    }
                    return LocalTime.parse(cell.toString());
                    
                case LOCAL_DATE_TIME:
                    if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                        return cell.getLocalDateTimeCellValue();
                    }
                    return LocalDateTime.parse(cell.toString());
                    
                default:
                    return cell.toString();
            }
        } catch (Exception e) {
            throw new ExcelIOException("Failed to convert cell value '" + cell.toString() + 
                "' to type " + targetType.getJavaClassName() + " at row " + cell.getRowIndex() + 
                ", column " + cell.getColumnIndex(), e);
        }
    }

    /**
     * Container for column metadata (names and types).
     */
    private static class ColumnMetadata {
        private final List<String> columnNames;
        private final List<ExcelJavaDataType> columnTypes;

        public ColumnMetadata(List<String> columnNames, List<ExcelJavaDataType> columnTypes) {
            if (columnNames.size() != columnTypes.size()) {
                throw new ExcelIOException("Column names and types must have the same size");
            }
            this.columnNames = columnNames;
            this.columnTypes = columnTypes;
        }

        public List<String> getColumnNames() {
            return columnNames;
        }

        public List<ExcelJavaDataType> getColumnTypes() {
            return columnTypes;
        }
    }
}