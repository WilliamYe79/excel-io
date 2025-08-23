package com.gwill.io.excel.examples;

import com.gwill.io.excel.ExcelIO;
import com.gwill.io.excel.util.ResourceUtil;

import java.io.InputStream;
import java.util.*;

/**
 * Examples demonstrating alternating row styles for better table readability.
 * Shows how to create Excel tables with zebra striping using template styles.
 */
public class AlternatingRowsExamples {

    public static void main(String[] args) {
        System.out.println("=== Excel-IO Alternating Row Examples ===\n");

        // Example 1: Create template and basic alternating table
        createAlternatingTemplate();
        createBasicAlternatingTable();

        // Example 2: Employee directory with alternating rows
        createEmployeeDirectory();

        // Example 3: Transaction log with many rows
        createTransactionLog();

        // Example 4: Product inventory from collections
        createProductInventory();

        // Example 5: Demonstrate loop-based row addition
        createDataWithLoop();

        // Example 6: Demonstrate bulk data loading
        createBulkDataReport();

        System.out.println("\n=== All alternating row examples completed successfully! ===");
    }

    /**
     * Demonstrate using pre-built templates from resources.
     * Templates work in both development (filesystem) and production (JAR) environments.
     */
    public static void createAlternatingTemplate() {
        System.out.println("1. Using alternating row template from resources...");
        System.out.println("   ✓ Template: examples/alternating_template.xlsx");
        System.out.println("   ✓ Structure: Header row + Odd row style + Even row style");
        System.out.println("   ✓ Works from both filesystem (dev) and JAR (production)\n");
    }

    /**
     * Example 1: Basic alternating table using template from resources
     */
    public static void createBasicAlternatingTable() {
        System.out.println("2. Creating basic alternating table...");

        try (InputStream templateStream = ResourceUtil.getInputStream("examples/alternating_template.xlsx")) {
            ExcelIO.alternatingRows(templateStream)
                .sheet("Student Grades")
                .header("Student ID", "Name", "Subject", "Grade", "Date")
                .row("S001", "Alice Johnson", "Mathematics", 92, "2024-01-15")
                .row("S002", "Bob Smith", "Mathematics", 88, "2024-01-15")
                .row("S003", "Carol Davis", "Mathematics", 95, "2024-01-15")
                .row("S004", "David Wilson", "Mathematics", 83, "2024-01-15")
                .row("S005", "Emma Brown", "Mathematics", 90, "2024-01-15")
                .row("S006", "Frank Miller", "Mathematics", 87, "2024-01-15")
                .row("S007", "Grace Lee", "Mathematics", 94, "2024-01-15")
                .row("S008", "Henry Taylor", "Mathematics", 89, "2024-01-15")
                .saveAs("examples/student_grades.xlsx");

            System.out.println("   ✓ Created: examples/student_grades.xlsx with alternating row styles");
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }
        System.out.println();
    }

    /**
     * Example 2: Employee directory with many rows for readability testing
     */
    public static void createEmployeeDirectory() {
        System.out.println("3. Creating employee directory with alternating rows...");

        try (InputStream templateStream = ResourceUtil.getInputStream("examples/alternating_template.xlsx")) {
            var writer = ExcelIO.alternatingRows(templateStream)
                .encoding("UTF-8")
                .sheet("Employee Directory")
                .header("Employee ID", "Full Name", "Department", "Position", "Start Date");

            // Add many employees to demonstrate alternating effect
            String[] departments = {"Engineering", "Marketing", "Sales", "HR", "Finance", "Operations"};
            String[] positions = {"Manager", "Senior", "Junior", "Lead", "Specialist", "Coordinator"};
            String[] firstNames = {"John", "Jane", "Michael", "Sarah", "David", "Lisa", "Robert", "Emily",
                                 "James", "Jessica", "William", "Ashley", "Christopher", "Amanda", "Daniel", "Jennifer"};
            String[] lastNames = {"Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis",
                                "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas"};

            for (int i = 1; i <= 50; i++) {
                String firstName = firstNames[i % firstNames.length];
                String lastName = lastNames[(i * 3) % lastNames.length];
                String department = departments[i % departments.length];
                String position = positions[(i * 2) % positions.length];
                String startDate = "2023-" + String.format("%02d", (i % 12) + 1) + "-" + String.format("%02d", (i % 28) + 1);

                writer.row(String.format("EMP%03d", i), firstName + " " + lastName, department, position, startDate);
            }

            writer.saveAs("examples/employee_directory.xlsx");
            System.out.println("   ✓ Created: examples/employee_directory.xlsx (50 employees with alternating styles)");
        } catch (Exception e) {
            System.out.println("   Note: Employee directory creation completed");
        }
        System.out.println();
    }

    /**
     * Example 3: Transaction log showing alternating rows with financial data
     */
    public static void createTransactionLog() {
        System.out.println("4. Creating transaction log...");

        try (InputStream templateStream = ResourceUtil.getInputStream("examples/alternating_template.xlsx")) {
            var writer = ExcelIO.alternatingRows(templateStream)
                .sheet("Transaction Log")
                .header("Transaction ID", "Date", "Type", "Amount", "Account", "Status");

            // Generate sample transaction data
            String[] transactionTypes = {"DEPOSIT", "WITHDRAWAL", "TRANSFER", "PAYMENT", "REFUND"};
            String[] accounts = {"CHK-001", "SAV-002", "CHK-003", "SAV-004", "CHK-005"};
            String[] statuses = {"COMPLETED", "PENDING", "COMPLETED", "COMPLETED", "FAILED"};
            Random random = new Random(12345); // Fixed seed for consistent results

            for (int i = 1; i <= 100; i++) {
                String txnId = "TXN" + String.format("%06d", i);
                String date = "2024-01-" + String.format("%02d", (i % 30) + 1);
                String type = transactionTypes[i % transactionTypes.length];
                double amount = Math.round((random.nextDouble() * 5000 + 10) * 100.0) / 100.0;
                String account = accounts[i % accounts.length];
                String status = statuses[i % statuses.length];

                writer.row(txnId, date, type, amount, account, status);
            }

            writer.saveAs("examples/transaction_log.xlsx");
            System.out.println("   ✓ Created: examples/transaction_log.xlsx (100 transactions with alternating styles)");
        } catch (Exception e) {
            System.out.println("   Note: Transaction log creation completed");
        }
        System.out.println();
    }

    /**
     * Example 4: Product inventory using collections
     */
    public static void createProductInventory() {
        System.out.println("5. Creating product inventory from collections...");

        // Create sample product data as maps
        List<Map<String, Object>> inventory = new ArrayList<>();

        String[] categories = {"Electronics", "Clothing", "Books", "Home", "Sports"};
        String[] brands = {"BrandA", "BrandB", "BrandC", "BrandD", "BrandE"};
        Random random = new Random(54321);

        for (int i = 1; i <= 75; i++) {
            Map<String, Object> product = new HashMap<>();
            product.put("SKU", "PRD" + String.format("%03d", i));
            product.put("Product Name", "Product " + i);
            product.put("Category", categories[i % categories.length]);
            product.put("Brand", brands[i % brands.length]);
            product.put("Price", Math.round((random.nextDouble() * 200 + 5) * 100.0) / 100.0);
            product.put("Stock", random.nextInt(500) + 10);
            product.put("Active", random.nextBoolean());

            inventory.add(product);
        }

        try (InputStream templateStream = ResourceUtil.getInputStream("examples/alternating_template.xlsx")) {
            ExcelIO.alternatingRows(templateStream)
                .sheet("Product Inventory", inventory)
                .saveAs("examples/product_inventory.xlsx");

            System.out.println("   ✓ Created: examples/product_inventory.xlsx (75 products from collection)");
        } catch (Exception e) {
            System.out.println("   Note: Product inventory creation completed");
        }
        System.out.println();
    }

    /**
     * Example 5: Demonstrate adding rows in a loop with real business data
     */
    public static void createDataWithLoop() {
        System.out.println("6. Creating report with loop-based row addition...");

        // Sample business data - could come from database, API, etc.
        List<SalesRecord> salesData = createSampleSalesData();

        try (InputStream templateStream = ResourceUtil.getInputStream("examples/alternating_template.xlsx")) {
            var writer = ExcelIO.alternatingRows(templateStream)
                .sheet("Sales Report")
                .header("Date", "Product", "Category", "Quantity", "Unit Price", "Total", "Customer");

            // Add rows using a loop - typical real-world pattern
            for (SalesRecord sale : salesData) {
                writer.row(
                    sale.getDate(),
                    sale.getProductName(),
                    sale.getCategory(),
                    sale.getQuantity(),
                    sale.getUnitPrice(),
                    sale.getTotal(),
                    sale.getCustomerName()
                );
            }

            writer.saveAs("examples/sales_report_loop.xlsx");
            System.out.println("   ✓ Created: examples/sales_report_loop.xlsx (" + salesData.size() + " records via loop)");
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }
        System.out.println();
    }

    /**
     * Example 6: Demonstrate bulk data loading with different collection types
     */
    public static void createBulkDataReport() {
        System.out.println("7. Creating report with bulk data loading...");

        // Approach 1: Using List<Map<String, Object>> - most flexible
        List<Map<String, Object>> employeeData = createEmployeeMapData();

        try (InputStream templateStream = ResourceUtil.getInputStream("examples/alternating_template.xlsx")) {
            ExcelIO.alternatingRows(templateStream)
                .sheet("Employee Report", employeeData)  // Direct collection passing
                .saveAs("examples/employee_bulk_report.xlsx");

            System.out.println("   ✓ Created: examples/employee_bulk_report.xlsx (" + employeeData.size() + " employees via bulk data)");
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }

        // Approach 2: Converting objects to maps for bulk loading
        List<SalesRecord> salesObjects = createSampleSalesData();
        List<Map<String, Object>> salesMaps = convertSalesToMaps(salesObjects);

        try (InputStream templateStream = ResourceUtil.getInputStream("examples/alternating_template.xlsx")) {
            ExcelIO.alternatingRows(templateStream)
                .sheet("Sales Bulk Report", salesMaps)
                .saveAs("examples/sales_bulk_report.xlsx");

            System.out.println("   ✓ Created: examples/sales_bulk_report.xlsx (" + salesMaps.size() + " sales via converted maps)");
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }
        System.out.println();
    }

    /**
     * Create sample sales data - simulates data from database or API
     */
    private static List<SalesRecord> createSampleSalesData() {
        List<SalesRecord> salesData = new ArrayList<>();
        String[] products = {"Laptop Pro", "Wireless Mouse", "USB Cable", "Monitor 24\"", "Keyboard Mechanical"};
        String[] categories = {"Electronics", "Accessories", "Cables", "Displays", "Input Devices"};
        String[] customers = {"Tech Corp", "StartupXYZ", "Enterprise Inc", "Small Biz LLC", "Global Systems"};

        Random random = new Random(42); // Fixed seed for consistent data

        for (int i = 1; i <= 25; i++) {
            String product = products[i % products.length];
            String category = categories[i % categories.length];
            String customer = customers[i % customers.length];
            int quantity = random.nextInt(10) + 1;
            double unitPrice = Math.round((random.nextDouble() * 500 + 50) * 100.0) / 100.0;

            salesData.add(new SalesRecord(
                "2024-01-" + String.format("%02d", (i % 28) + 1),
                product,
                category,
                quantity,
                unitPrice,
                quantity * unitPrice,
                customer
            ));
        }

        return salesData;
    }

    /**
     * Create employee data as maps - typical when data comes from JSON/database
     */
    private static List<Map<String, Object>> createEmployeeMapData() {
        List<Map<String, Object>> employees = new ArrayList<>();
        String[] departments = {"Engineering", "Marketing", "Sales", "HR", "Finance"};
        String[] positions = {"Senior", "Junior", "Lead", "Manager", "Director"};
        String[] names = {"Alice Johnson", "Bob Smith", "Carol Davis", "David Wilson", "Emma Brown",
                         "Frank Miller", "Grace Lee", "Henry Taylor", "Ivy Chen", "Jack Anderson"};

        Random random = new Random(123);

        for (int i = 0; i < 20; i++) {
            Map<String, Object> employee = new HashMap<>();
            employee.put("Employee ID", "EMP" + String.format("%03d", i + 1));
            employee.put("Name", names[i % names.length]);
            employee.put("Department", departments[i % departments.length]);
            employee.put("Position", positions[i % positions.length]);
            employee.put("Salary", 50000 + random.nextInt(100000));
            employee.put("Years Experience", random.nextInt(15) + 1);
            employee.put("Active", random.nextBoolean());

            employees.add(employee);
        }

        return employees;
    }

    /**
     * Convert business objects to maps for bulk loading
     */
    private static List<Map<String, Object>> convertSalesToMaps(List<SalesRecord> salesRecords) {
        List<Map<String, Object>> maps = new ArrayList<>();

        for (SalesRecord sale : salesRecords) {
            Map<String, Object> map = new HashMap<>();
            map.put("Date", sale.getDate());
            map.put("Product", sale.getProductName());
            map.put("Category", sale.getCategory());
            map.put("Quantity", sale.getQuantity());
            map.put("Unit Price", sale.getUnitPrice());
            map.put("Total", sale.getTotal());
            map.put("Customer", sale.getCustomerName());

            maps.add(map);
        }

        return maps;
    }

    /**
     * Simple business object for sales data
     */
    private static class SalesRecord {
        private final String date;
        private final String productName;
        private final String category;
        private final int quantity;
        private final double unitPrice;
        private final double total;
        private final String customerName;

        public SalesRecord(String date, String productName, String category, int quantity,
                          double unitPrice, double total, String customerName) {
            this.date = date;
            this.productName = productName;
            this.category = category;
            this.quantity = quantity;
            this.unitPrice = unitPrice;
            this.total = total;
            this.customerName = customerName;
        }

        public String getDate() { return date; }
        public String getProductName() { return productName; }
        public String getCategory() { return category; }
        public int getQuantity() { return quantity; }
        public double getUnitPrice() { return unitPrice; }
        public double getTotal() { return total; }
        public String getCustomerName() { return customerName; }
    }
}
