package com.gwill.io.excel.examples;

import com.gwill.io.excel.ExcelIO;

import java.time.LocalDate;
import java.util.*;

/**
 * Basic usage examples for excel-io library.
 * Demonstrates simple reading and writing operations.
 */
public class BasicUsageExamples {

    public static void main(String[] args) {
        System.out.println("=== Excel-IO Basic Usage Examples ===\n");
        
        // Example 1: Simple Writing
        createBasicExcelFile();
        
        // Example 2: Reading Data  
        readBasicExcelFile();
        
        // Example 3: Writing from Collections
        writeFromCollections();
        
        // Example 4: Multiple Sheets
        createMultiSheetWorkbook();
        
        // Example 5: International Characters
        createInternationalExcel();
        
        // Example 6: Bulk data operations
        demonstrateBulkOperations();
        
        System.out.println("\n=== All basic examples completed successfully! ===");
    }

    /**
     * Example 1: Create a simple Excel file with headers and data rows
     */
    public static void createBasicExcelFile() {
        System.out.println("1. Creating basic Excel file...");
        
        ExcelIO.write("examples/basic_example.xlsx")
            .sheet("Employee Data")
            .header("Employee ID", "Name", "Department", "Salary", "Start Date")
            .row("EMP001", "John Doe", "Engineering", 75000, "2023-01-15")
            .row("EMP002", "Jane Smith", "Marketing", 65000, "2023-02-01")  
            .row("EMP003", "Mike Johnson", "Sales", 58000, "2023-01-20")
            .row("EMP004", "Sarah Wilson", "HR", 62000, "2023-03-01")
            .save();
            
        System.out.println("   ✓ Created: examples/basic_example.xlsx\n");
    }

    /**
     * Example 2: Read data from Excel file
     */
    public static void readBasicExcelFile() {
        System.out.println("2. Reading Excel file...");
        
        try {
            List<Map<String, Object>> employees = ExcelIO.read("examples/basic_example.xlsx")
                .sheet("Employee Data")
                .asMaps();
                
            System.out.println("   Read " + employees.size() + " employee records:");
            for (Map<String, Object> employee : employees) {
                System.out.println("   - " + employee.get("Name") + " (" + employee.get("Department") + "): $" + employee.get("Salary"));
            }
        } catch (Exception e) {
            System.out.println("   Note: Could not read file (will be created by previous example)");
        }
        System.out.println();
    }

    /**
     * Example 3: Write data from collections  
     */
    public static void writeFromCollections() {
        System.out.println("3. Writing from collections...");
        
        // Create sample data
        List<Map<String, Object>> products = Arrays.asList(
            Map.of("SKU", "LAPTOP001", "Name", "Dell XPS 13", "Price", 1299.99, "InStock", true),
            Map.of("SKU", "PHONE002", "Name", "iPhone 15", "Price", 999.99, "InStock", false),
            Map.of("SKU", "TABLET003", "Name", "iPad Pro", "Price", 1099.99, "InStock", true)
        );
        
        ExcelIO.write("examples/products.xlsx")
            .sheet("Products", products)
            .save();
            
        System.out.println("   ✓ Created: examples/products.xlsx from collection\n");
    }

    /**
     * Example 4: Create workbook with multiple sheets
     */
    public static void createMultiSheetWorkbook() {
        System.out.println("4. Creating multi-sheet workbook...");
        
        ExcelIO.write("examples/company_report.xlsx")
            // Sheet 1: Employees
            .sheet("Employees")
            .header("ID", "Name", "Position", "Salary")
            .row("E001", "Alice Brown", "Manager", 85000)
            .row("E002", "Bob Davis", "Developer", 70000)
            .row("E003", "Carol White", "Designer", 68000)
            
            // Sheet 2: Departments
            .sheet("Departments")  
            .header("Dept Code", "Department", "Head", "Budget")
            .row("ENG", "Engineering", "Alice Brown", 500000)
            .row("DES", "Design", "Carol White", 300000)
            .row("HR", "Human Resources", "David Green", 200000)
            
            // Sheet 3: Projects
            .sheet("Projects")
            .header("Project", "Status", "Budget", "Deadline")
            .row("Website Redesign", "In Progress", 150000, "2024-06-30")
            .row("Mobile App", "Planning", 200000, "2024-12-15")
            .save();
            
        System.out.println("   ✓ Created: examples/company_report.xlsx with 3 sheets\n");
    }

    /**
     * Example 5: Handle international characters with proper encoding
     */
    public static void createInternationalExcel() {
        System.out.println("5. Creating Excel with international characters...");
        
        ExcelIO.write("examples/international.xlsx")
            .encoding("UTF-8")
            .sheet("多语言数据")  // Chinese sheet name
            .header("姓名", "城市", "国家", "工资", "备注")  // Chinese headers
            .row("张三", "北京", "中国", 50000, "软件工程师")
            .row("José García", "Madrid", "España", 45000, "Desarrollador")
            .row("François Dubois", "Paris", "France", 48000, "Développeur")
            .row("田中太郎", "東京", "日本", 55000, "プログラマー")
            .save();
            
        System.out.println("   ✓ Created: examples/international.xlsx with UTF-8 encoding\n");
    }
    
    /**
     * Example 6: Demonstrate bulk data operations for real-world scenarios
     */
    public static void demonstrateBulkOperations() {
        System.out.println("6. Demonstrating bulk data operations...");
        
        // Create large dataset for bulk operations
        List<Map<String, Object>> employeeData = createLargeEmployeeDataset();
        
        // Write bulk data to Excel
        ExcelIO.write("examples/bulk_employee_data.xlsx")
            .sheet("All Employees", employeeData)
            .save();
            
        System.out.println("   ✓ Created: examples/bulk_employee_data.xlsx (" + employeeData.size() + " employees)");
        
        // Read back and process the data
        try {
            List<Map<String, Object>> readData = ExcelIO.read("examples/bulk_employee_data.xlsx")
                .sheet("All Employees")
                .asMaps();
                
            System.out.println("   ✓ Read back " + readData.size() + " employee records");
            
            // Demonstrate data filtering and processing
            long seniorEmployees = readData.stream()
                .filter(emp -> {
                    Object yearsObj = emp.get("Years Experience");
                    if (yearsObj instanceof Number) {
                        return ((Number) yearsObj).intValue() >= 10;
                    }
                    return false;
                })
                .count();
                
            System.out.println("   ✓ Found " + seniorEmployees + " senior employees (10+ years experience)");
            
        } catch (Exception e) {
            System.out.println("   Error reading bulk data: " + e.getMessage());
        }
        
        System.out.println();
    }
    
    /**
     * Create a large employee dataset for bulk operations demo
     */
    private static List<Map<String, Object>> createLargeEmployeeDataset() {
        List<Map<String, Object>> employees = new ArrayList<>();
        String[] departments = {"Engineering", "Marketing", "Sales", "HR", "Finance", "Operations", "Support"};
        String[] positions = {"Junior", "Senior", "Lead", "Manager", "Director", "VP"};
        String[] firstNames = {"Alice", "Bob", "Carol", "David", "Emma", "Frank", "Grace", "Henry", 
                              "Ivy", "Jack", "Kate", "Liam", "Mia", "Noah", "Olivia", "Peter"};
        String[] lastNames = {"Anderson", "Brown", "Davis", "Garcia", "Johnson", "Jones", "Miller", 
                              "Smith", "Taylor", "Wilson", "Martinez", "Lopez", "Gonzalez", "Rodriguez"};
        
        Random random = new Random(999);
        
        for (int i = 1; i <= 150; i++) {
            Map<String, Object> employee = new HashMap<>();
            
            String firstName = firstNames[i % firstNames.length];
            String lastName = lastNames[(i * 3) % lastNames.length];
            
            employee.put("Employee ID", "EMP" + String.format("%04d", i));
            employee.put("Full Name", firstName + " " + lastName);
            employee.put("Department", departments[i % departments.length]);
            employee.put("Position", positions[i % positions.length]);
            employee.put("Salary", 45000 + random.nextInt(120000));
            employee.put("Years Experience", random.nextInt(25) + 1);
            employee.put("Active", random.nextBoolean());
            employee.put("Start Date", "202" + (random.nextInt(4) + 0) + "-" + 
                        String.format("%02d", random.nextInt(12) + 1) + "-" + 
                        String.format("%02d", random.nextInt(28) + 1));
            
            employees.add(employee);
        }
        
        return employees;
    }
}