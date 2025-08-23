package com.gwill.io.excel;

import java.io.File;
import java.io.InputStream;
import java.nio.file.Path;

/**
 * Main entry point for Excel I/O operations.
 * Provides a fluent, intuitive API for reading and writing Excel files.
 * 
 * <h3>Quick Start - Writing</h3>
 * <pre>{@code
 * ExcelIO.write("output.xlsx")
 *     .sheet("Users")
 *     .header("Name", "Age", "Email")
 *     .row("John Doe", 30, "john@example.com")
 *     .row("Jane Smith", 25, "jane@example.com")
 *     .save();
 * }</pre>
 * 
 * <h3>Quick Start - Reading</h3>
 * <pre>{@code
 * List<Map<String, Object>> data = ExcelIO.read("input.xlsx")
 *     .sheet("Users")
 *     .asMaps();
 * }</pre>
 * 
 * @author William Ye
 * @version 1.0.0
 */
public final class ExcelIO {

    private ExcelIO() {
        // Utility class - prevent instantiation
    }

    /**
     * Start building an Excel file for writing.
     * 
     * @param filePath the path where the Excel file will be saved
     * @return a new {@link ExcelWriter} builder
     */
    public static ExcelWriter write(String filePath) {
        return new ExcelWriter(filePath);
    }

    /**
     * Start building a categorized Excel file for writing.
     * Categorized files support summary rows with detail rows underneath.
     * 
     * @param filePath the path where the Excel file will be saved
     * @return a new {@link CategorizedWriter} builder
     */
    public static CategorizedWriter writeCategorized(String filePath) {
        return new CategorizedWriter(filePath);
    }

    /**
     * Start building a categorized Excel file for writing.
     * Categorized files support summary rows with detail rows underneath.
     * 
     * @param file the File where the Excel file will be saved
     * @return a new {@link CategorizedWriter} builder
     */
    public static CategorizedWriter writeCategorized(File file) {
        return new CategorizedWriter(file.getAbsolutePath());
    }

    /**
     * Start building a categorized Excel file for writing.
     * Categorized files support summary rows with detail rows underneath.
     * 
     * @param path the Path where the Excel file will be saved
     * @return a new {@link CategorizedWriter} builder
     */
    public static CategorizedWriter writeCategorized(Path path) {
        return new CategorizedWriter(path.toString());
    }

    /**
     * Start building an Excel file for writing.
     * 
     * @param file the File where the Excel file will be saved
     * @return a new {@link ExcelWriter} builder
     */
    public static ExcelWriter write(File file) {
        return new ExcelWriter(file.getAbsolutePath());
    }

    /**
     * Start building an Excel file for writing.
     * 
     * @param path the Path where the Excel file will be saved
     * @return a new {@link ExcelWriter} builder
     */
    public static ExcelWriter write(Path path) {
        return new ExcelWriter(path.toString());
    }

    /**
     * Start reading from an Excel file.
     * 
     * @param filePath the path to the Excel file to read
     * @return a new {@link ExcelReader} for the file
     */
    public static ExcelReader read(String filePath) {
        return new ExcelReader(filePath);
    }

    /**
     * Start reading from an Excel file.
     * 
     * @param file the Excel file to read
     * @return a new {@link ExcelReader} for the file
     */
    public static ExcelReader read(File file) {
        return new ExcelReader(file.getAbsolutePath());
    }

    /**
     * Start reading from an Excel file.
     * 
     * @param path the Path to the Excel file to read
     * @return a new {@link ExcelReader} for the file
     */
    public static ExcelReader read(Path path) {
        return new ExcelReader(path.toString());
    }

    /**
     * Start reading from an Excel InputStream.
     * 
     * @param inputStream the InputStream containing Excel data
     * @return a new {@link ExcelReader} for the stream
     */
    public static ExcelReader read(InputStream inputStream) {
        return new ExcelReader(inputStream);
    }

    /**
     * Start reading from an Excel file with type metadata.
     * Requires metadata specification via withMetadata() or withTypes().
     * 
     * @param filePath the path to the Excel file to read
     * @return a new {@link ExcelReader} for typed reading
     * @deprecated Use {@link #read(String)} and call {@code withMetadata()} or {@code withTypes()} instead
     */
    @Deprecated
    public static ExcelReader readTyped(String filePath) {
        return new ExcelReader(filePath);
    }

    /**
     * Start reading from an Excel file with type metadata.
     * Requires metadata specification via withMetadata() or withTypes().
     * 
     * @param file the Excel file to read
     * @return a new {@link ExcelReader} for typed reading
     * @deprecated Use {@link #read(File)} and call {@code withMetadata()} or {@code withTypes()} instead
     */
    @Deprecated
    public static ExcelReader readTyped(File file) {
        return new ExcelReader(file.getAbsolutePath());
    }

    /**
     * Start reading from an Excel file with type metadata.
     * Requires metadata specification via withMetadata() or withTypes().
     * 
     * @param path the Path to the Excel file to read
     * @return a new {@link ExcelReader} for typed reading
     * @deprecated Use {@link #read(Path)} and call {@code withMetadata()} or {@code withTypes()} instead
     */
    @Deprecated
    public static ExcelReader readTyped(Path path) {
        return new ExcelReader(path.toString());
    }

    /**
     * Start reading from an Excel InputStream with type metadata.
     * Requires metadata specification via withMetadata() or withTypes().
     * 
     * @param inputStream the InputStream containing Excel data
     * @return a new {@link ExcelReader} for typed reading
     * @deprecated Use {@link #read(InputStream)} and call {@code withMetadata()} or {@code withTypes()} instead
     */
    @Deprecated
    public static ExcelReader readTyped(InputStream inputStream) {
        return new ExcelReader(inputStream);
    }

    /**
     * Create a form-based Excel file from a template.
     * Perfect for business documents like invoices, purchase orders, and reports.
     * The template should follow a 16-row structure with form headers, table section, and footers.
     * 
     * @param templatePath the path to the template Excel file
     * @return a new {@link FormTemplateWriter} for form-based template writing
     */
    public static FormTemplateWriter formTemplate(String templatePath) {
        return new FormTemplateWriter(templatePath);
    }

    /**
     * Create a form-based Excel file from a template.
     * Perfect for business documents like invoices, purchase orders, and reports.
     * The template should follow a 16-row structure with form headers, table section, and footers.
     * 
     * @param templateFile the template Excel file
     * @return a new {@link FormTemplateWriter} for form-based template writing
     */
    public static FormTemplateWriter formTemplate(File templateFile) {
        return new FormTemplateWriter(templateFile.getAbsolutePath());
    }

    /**
     * Create a form-based Excel file from a template InputStream.
     * Perfect for business documents like invoices, purchase orders, and reports.
     * The template should follow a 16-row structure with form headers, table section, and footers.
     * 
     * @param templateStream the InputStream containing template Excel data
     * @return a new {@link FormTemplateWriter} for form-based template writing
     */
    public static FormTemplateWriter formTemplate(InputStream templateStream) {
        return new FormTemplateWriter(templateStream);
    }

    /**
     * Create a categorized Excel file from a template with predefined styles.
     * The template should have 3 rows with the desired styles:
     * Row 0: Header styles, Row 1: Category styles, Row 2: Detail styles.
     * 
     * @param templatePath the path to the template Excel file
     * @return a new {@link CategorizedTemplateWriter} for template-based categorized writing
     */
    public static CategorizedTemplateWriter categorizedFromTemplate(String templatePath) {
        return new CategorizedTemplateWriter(templatePath);
    }

    /**
     * Create a categorized Excel file from a template with predefined styles.
     * The template should have 3 rows with the desired styles:
     * Row 0: Header styles, Row 1: Category styles, Row 2: Detail styles.
     * 
     * @param templateFile the template Excel file
     * @return a new {@link CategorizedTemplateWriter} for template-based categorized writing
     */
    public static CategorizedTemplateWriter categorizedFromTemplate(File templateFile) {
        return new CategorizedTemplateWriter(templateFile.getAbsolutePath());
    }

    /**
     * Create a categorized Excel file from a template InputStream with predefined styles.
     * The template should have 3 rows with the desired styles:
     * Row 0: Header styles, Row 1: Category styles, Row 2: Detail styles.
     * 
     * @param templateStream the InputStream containing template Excel data
     * @return a new {@link CategorizedTemplateWriter} for template-based categorized writing
     */
    public static CategorizedTemplateWriter categorizedFromTemplate(InputStream templateStream) {
        return new CategorizedTemplateWriter(templateStream);
    }

    /**
     * Create an Excel file with alternating row styles from a template.
     * The template should have 3 rows with the desired styles:
     * Row 0: Header styles, Row 1: Odd row styles, Row 2: Even row styles.
     * 
     * @param templatePath the path to the template Excel file
     * @return a new {@link AlternatingRowsWriter} for template-based alternating row writing
     */
    public static AlternatingRowsWriter alternatingRows(String templatePath) {
        return new AlternatingRowsWriter(templatePath);
    }

    /**
     * Create an Excel file with alternating row styles from a template.
     * The template should have 3 rows with the desired styles:
     * Row 0: Header styles, Row 1: Odd row styles, Row 2: Even row styles.
     * 
     * @param templateFile the template Excel file
     * @return a new {@link AlternatingRowsWriter} for template-based alternating row writing
     */
    public static AlternatingRowsWriter alternatingRows(File templateFile) {
        return new AlternatingRowsWriter(templateFile.getAbsolutePath());
    }

    /**
     * Create an Excel file with alternating row styles from a template InputStream.
     * The template should have 3 rows with the desired styles:
     * Row 0: Header styles, Row 1: Odd row styles, Row 2: Even row styles.
     * 
     * @param templateStream the InputStream containing template Excel data
     * @return a new {@link AlternatingRowsWriter} for template-based alternating row writing
     */
    public static AlternatingRowsWriter alternatingRows(InputStream templateStream) {
        return new AlternatingRowsWriter(templateStream);
    }
}