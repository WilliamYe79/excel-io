package com.gwill.io.excel.examples;

import com.gwill.io.excel.ExcelIO;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.stream.Collectors;

/**
 * Examples demonstrating type-safe Excel reading with metadata.
 * Shows how to read Excel files with proper data type conversion.
 */
public class TypedReadingExamples {

    public static void main(String[] args) {
        System.out.println("=== Excel-IO Typed Reading Examples ===\n");
        
        // Example 1: Create sample data and metadata
        createSampleDataAndMetadata();
        
        // Example 2: Read with metadata file
        readWithMetadataFile();
        
        // Example 3: Read with inline type definitions
        readWithInlineTypes();
        
        // Example 4: Handle different data types
        demonstrateDataTypes();
        
        // Example 5: Demonstrate reading from bulk data collections
        demonstrateBulkDataReading();
        
        System.out.println("\n=== All typed reading examples completed successfully! ===");
    }

    /**
     * Example 1: Create sample data and its metadata file
     */
    public static void createSampleDataAndMetadata() {
        System.out.println("1. Creating sample data and metadata files...");
        
        // Create sample financial data
        ExcelIO.write("examples/financial_data.xlsx")
            .sheet("Transactions")
            .header("Transaction ID", "Date", "Amount", "Tax Rate", "Is Processed", "Description")
            .row("TXN001", "2024-01-15", 1250.75, 0.08, true, "Equipment Purchase")
            .row("TXN002", "2024-01-16", 850.50, 0.08, false, "Software License")
            .row("TXN003", "2024-01-17", 2100.25, 0.10, true, "Consulting Services")
            .row("TXN004", "2024-01-18", 675.00, 0.08, false, "Office Supplies")
            .save();
        
        // Create metadata file defining data types
        ExcelIO.write("examples/financial_metadata.xlsx")
            .sheet("Metadata")
            .header("Transaction ID", "Date", "Amount", "Tax Rate", "Is Processed", "Description")
            .row("java.lang.String", "java.time.LocalDate", "java.math.BigDecimal", 
                 "java.lang.Double", "java.lang.Boolean", "java.lang.String")
            .save();
            
        System.out.println("   ✓ Created sample data: examples/financial_data.xlsx");
        System.out.println("   ✓ Created metadata: examples/financial_metadata.xlsx\n");
    }

    /**
     * Example 2: Read Excel data using metadata file for type conversion
     */
    public static void readWithMetadataFile() {
        System.out.println("2. Reading with metadata file...");
        
        try {
            List<Map<String, Object>> transactions = ExcelIO.read("examples/financial_data.xlsx")
                .withMetadata("examples/financial_metadata.xlsx")
                .sheet("Transactions")
                .asMaps();
                
            System.out.println("   Read " + transactions.size() + " transactions with proper types:");
            
            for (Map<String, Object> transaction : transactions) {
                String id = (String) transaction.get("Transaction ID");
                LocalDate date = (LocalDate) transaction.get("Date");
                BigDecimal amount = (BigDecimal) transaction.get("Amount");
                Double taxRate = (Double) transaction.get("Tax Rate");
                Boolean isProcessed = (Boolean) transaction.get("Is Processed");
                String description = (String) transaction.get("Description");
                
                System.out.printf("   - %s: $%s on %s (Tax: %.1f%%) [%s] - %s%n",
                    id, amount, date, taxRate * 100, 
                    isProcessed ? "Processed" : "Pending", description);
            }
        } catch (Exception e) {
            System.out.println("   Note: Could not read files (will be created by previous example)");
        }
        System.out.println();
    }

    /**
     * Example 3: Read with inline type definitions (no metadata file needed)
     */
    public static void readWithInlineTypes() {
        System.out.println("3. Reading with inline type definitions...");
        
        // Create simple employee data
        ExcelIO.write("examples/employees_simple.xlsx")
            .sheet("Staff")
            .header("ID", "Name", "Age", "Salary", "Active")
            .row(1001, "John Doe", 30, 75000.50, true)
            .row(1002, "Jane Smith", 28, 68000.75, true) 
            .row(1003, "Mike Wilson", 35, 82000.00, false)
            .save();
            
        try {
            // Read with inline type specification
            List<Map<String, Object>> employees = ExcelIO.read("examples/employees_simple.xlsx")
                .withTypes("ID:java.lang.Integer", 
                          "Name:java.lang.String",
                          "Age:java.lang.Integer", 
                          "Salary:java.lang.Double",
                          "Active:java.lang.Boolean")
                .sheet("Staff")
                .asMaps();
                
            System.out.println("   Read " + employees.size() + " employees with inline types:");
            
            for (Map<String, Object> emp : employees) {
                Integer id = (Integer) emp.get("ID");
                String name = (String) emp.get("Name");
                Integer age = (Integer) emp.get("Age");
                Double salary = (Double) emp.get("Salary");
                Boolean active = (Boolean) emp.get("Active");
                
                System.out.printf("   - ID:%d %s (Age:%d) $%.2f [%s]%n",
                    id, name, age, salary, active ? "Active" : "Inactive");
            }
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }
        System.out.println();
    }

    /**
     * Example 4: Demonstrate various supported data types
     */
    public static void demonstrateDataTypes() {
        System.out.println("4. Demonstrating various data types...");
        
        // Create data with various types
        ExcelIO.write("examples/data_types_demo.xlsx")
            .sheet("Types Demo")
            .header("Text", "WholeNumber", "Decimal", "Money", "Flag", "Birthday", "UpdatedAt")
            .row("Sample Text", 42, 3.14159, "1250.75", true, "1990-05-15", "2024-01-15")
            .row("另一个样本", 100, 2.71828, "2500.50", false, "1985-12-25", "2024-01-16")
            .save();
            
        try {
            List<Map<String, Object>> records = ExcelIO.read("examples/data_types_demo.xlsx")
                .withTypes("Text:java.lang.String",
                          "WholeNumber:java.lang.Integer", 
                          "Decimal:java.lang.Double",
                          "Money:java.math.BigDecimal",
                          "Flag:java.lang.Boolean",
                          "Birthday:java.time.LocalDate",
                          "UpdatedAt:java.time.LocalDate")
                .sheet("Types Demo")
                .asMaps();
                
            System.out.println("   Demonstrating type conversion:");
            
            for (Map<String, Object> record : records) {
                System.out.println("   - Text: " + record.get("Text") + " (" + record.get("Text").getClass().getSimpleName() + ")");
                System.out.println("     Number: " + record.get("WholeNumber") + " (" + record.get("WholeNumber").getClass().getSimpleName() + ")");
                System.out.println("     Decimal: " + record.get("Decimal") + " (" + record.get("Decimal").getClass().getSimpleName() + ")");
                System.out.println("     Money: " + record.get("Money") + " (" + record.get("Money").getClass().getSimpleName() + ")");
                System.out.println("     Boolean: " + record.get("Flag") + " (" + record.get("Flag").getClass().getSimpleName() + ")");
                System.out.println("     Date: " + record.get("Birthday") + " (" + record.get("Birthday").getClass().getSimpleName() + ")");
                System.out.println();
            }
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }
    }
    
    /**
     * Example 5: Demonstrate reading and processing bulk data collections
     */
    public static void demonstrateBulkDataReading() {
        System.out.println("5. Demonstrating bulk data reading and processing...");
        
        // Create a larger dataset for bulk processing demonstration
        createLargeDataset();
        
        try {
            // Read large dataset with proper types
            List<Map<String, Object>> orders = ExcelIO.read("examples/large_orders.xlsx")
                .withTypes("Order ID:java.lang.String",
                          "Customer:java.lang.String", 
                          "Product:java.lang.String",
                          "Quantity:java.lang.Integer",
                          "Price:java.math.BigDecimal",
                          "Order Date:java.time.LocalDate",
                          "Shipped:java.lang.Boolean")
                .sheet("Orders")
                .asMaps();
                
            System.out.println("   Read " + orders.size() + " orders with proper type conversion");
            
            // Demonstrate data processing with proper types
            processOrderData(orders);
            
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }
        System.out.println();
    }
    
    /**
     * Create a larger dataset for bulk reading demonstration
     */
    private static void createLargeDataset() {
        var writer = ExcelIO.write("examples/large_orders.xlsx")
            .sheet("Orders")
            .header("Order ID", "Customer", "Product", "Quantity", "Price", "Order Date", "Shipped");
            
        String[] customers = {"TechCorp Inc", "StartupXYZ", "Enterprise Ltd", "SmallBiz Co", "GlobalSys", 
                             "InnovateLab", "CloudFirst", "DataDriven", "AgileTeam", "ScaleUp LLC"};
        String[] products = {"Software License", "Hardware Kit", "Consulting Hours", "Training Package", 
                           "Support Contract", "Cloud Credits", "Development Tools", "Security Suite"};
        Random random = new Random(456);
        
        // Generate 100 orders for bulk processing demo
        for (int i = 1; i <= 100; i++) {
            String orderId = "ORD" + String.format("%04d", i);
            String customer = customers[i % customers.length];
            String product = products[i % products.length];
            int quantity = random.nextInt(10) + 1;
            double price = Math.round((random.nextDouble() * 1000 + 50) * 100.0) / 100.0;
            String orderDate = "2024-" + String.format("%02d", (i % 12) + 1) + "-" + String.format("%02d", (i % 28) + 1);
            boolean shipped = random.nextBoolean();
            
            writer.row(orderId, customer, product, quantity, price, orderDate, shipped);
        }
        
        writer.save();
    }
    
    /**
     * Process order data to demonstrate type-safe operations
     */
    private static void processOrderData(List<Map<String, Object>> orders) {
        System.out.println("   Processing orders with type-safe operations:");
        
        // Calculate total revenue (BigDecimal arithmetic)
        BigDecimal totalRevenue = orders.stream()
            .map(order -> {
                Integer quantity = (Integer) order.get("Quantity");
                BigDecimal price = (BigDecimal) order.get("Price");
                return price.multiply(BigDecimal.valueOf(quantity));
            })
            .reduce(BigDecimal.ZERO, BigDecimal::add);
            
        System.out.println("     - Total Revenue: $" + totalRevenue);
        
        // Count shipped vs unshipped orders (Boolean operations)
        long shippedCount = orders.stream()
            .mapToLong(order -> (Boolean) order.get("Shipped") ? 1 : 0)
            .sum();
            
        System.out.println("     - Shipped Orders: " + shippedCount + "/" + orders.size());
        
        // Find orders from current month (LocalDate operations)
        LocalDate now = LocalDate.now();
        long currentMonthOrders = orders.stream()
            .filter(order -> {
                LocalDate orderDate = (LocalDate) order.get("Order Date");
                return orderDate.getYear() == now.getYear() && orderDate.getMonth() == now.getMonth();
            })
            .count();
            
        System.out.println("     - Current Month Orders: " + currentMonthOrders);
        
        // Group by customer (String operations)
        Map<String, Long> customerOrderCounts = orders.stream()
            .collect(Collectors.groupingBy(
                order -> (String) order.get("Customer"),
                Collectors.counting()
            ));
            
        System.out.println("     - Top Customer: " + 
            customerOrderCounts.entrySet().stream()
                .max(Map.Entry.comparingByValue())
                .map(entry -> entry.getKey() + " (" + entry.getValue() + " orders)")
                .orElse("None"));
    }
}