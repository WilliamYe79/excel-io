package com.gwill.io.excel.examples;

import com.gwill.io.excel.ExcelIO;
import com.gwill.io.excel.util.ResourceUtil;

import java.io.InputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Real-world example: E-commerce Sales Analysis System
 * 
 * This example demonstrates a complete workflow:
 * 1. Read raw sales data with proper type conversion
 * 2. Process and analyze the data
 * 3. Generate categorized reports with template styling
 * 4. Create alternating row tables for large datasets
 * 
 * Similar to your Amazon ASIN performance analysis system.
 */
public class RealWorldExample {

    public static void main(String[] args) {
        System.out.println("=== Real-World Example: E-commerce Sales Analysis ===\n");
        
        // Step 1: Create sample raw data and metadata
        createRawSalesData();
        
        // Step 2: Read and process the data
        List<SalesRecord> salesData = readAndProcessSalesData();
        
        // Step 3: Generate categorized performance report
        generatePerformanceReport(salesData);
        
        // Step 4: Generate detailed transaction list with alternating rows
        generateDetailedTransactionReport(salesData);
        
        // Step 5: Generate executive summary
        generateExecutiveSummary(salesData);
        
        System.out.println("\n=== E-commerce analysis completed successfully! ===");
    }

    /**
     * Step 1: Create sample raw sales data and its metadata
     */
    public static void createRawSalesData() {
        System.out.println("1. Creating raw sales data...");
        
        // Create raw sales data (simulating export from e-commerce system)
        // In real applications, this data would come from database/API
        ExcelIO.write("examples/raw_sales_data.xlsx")
            .sheet("Sales Export")
            .header("Order ID", "Product SKU", "Product Name", "Category", "Sale Date", 
                   "Quantity", "Unit Price", "Total Amount", "Customer Segment", "Is Premium")
            
            // Electronics - High performers
            .row("ORD001", "ELEC001", "Wireless Headphones Pro", "Electronics", "2024-01-15", 2, 199.99, 399.98, "Premium", true)
            .row("ORD002", "ELEC002", "Smart Watch Series X", "Electronics", "2024-01-16", 1, 349.99, 349.99, "Premium", true)
            .row("ORD003", "ELEC001", "Wireless Headphones Pro", "Electronics", "2024-01-17", 1, 199.99, 199.99, "Regular", false)
            .row("ORD004", "ELEC003", "Bluetooth Speaker", "Electronics", "2024-01-18", 3, 89.99, 269.97, "Regular", false)
            .row("ORD011", "ELEC001", "Wireless Headphones Pro", "Electronics", "2024-01-22", 3, 199.99, 599.97, "Premium", true)
            .row("ORD014", "ELEC002", "Smart Watch Series X", "Electronics", "2024-01-25", 2, 349.99, 699.98, "Premium", true)
            
            // Fashion - Medium performers
            .row("ORD005", "FASH001", "Designer Jacket", "Fashion", "2024-01-16", 1, 299.99, 299.99, "Premium", true)
            .row("ORD006", "FASH002", "Running Shoes", "Fashion", "2024-01-17", 2, 149.99, 299.98, "Regular", false)
            .row("ORD007", "FASH003", "Casual T-Shirt", "Fashion", "2024-01-18", 5, 29.99, 149.95, "Regular", false)
            .row("ORD012", "FASH001", "Designer Jacket", "Fashion", "2024-01-23", 1, 299.99, 299.99, "Premium", true)
            .row("ORD015", "FASH002", "Running Shoes", "Fashion", "2024-01-26", 1, 149.99, 149.99, "Regular", false)
            
            // Home & Garden - Lower performers
            .row("ORD008", "HOME001", "Coffee Maker Deluxe", "Home", "2024-01-19", 1, 179.99, 179.99, "Premium", true)
            .row("ORD009", "HOME002", "Plant Pot Set", "Home", "2024-01-20", 2, 24.99, 49.98, "Regular", false)
            .row("ORD010", "HOME003", "LED Desk Lamp", "Home", "2024-01-21", 1, 79.99, 79.99, "Regular", false)
            .row("ORD013", "HOME001", "Coffee Maker Deluxe", "Home", "2024-01-24", 2, 179.99, 359.98, "Premium", true)
            
            .save();
        
        // Create metadata for proper type conversion
        ExcelIO.write("examples/sales_metadata.xlsx")
            .sheet("Column Types")
            .header("Order ID", "Product SKU", "Product Name", "Category", "Sale Date", 
                   "Quantity", "Unit Price", "Total Amount", "Customer Segment", "Is Premium")
            .row("java.lang.String", "java.lang.String", "java.lang.String", "java.lang.String", "java.time.LocalDate",
                 "java.lang.Integer", "java.math.BigDecimal", "java.math.BigDecimal", "java.lang.String", "java.lang.Boolean")
            .save();
            
        System.out.println("   ✓ Created raw sales data with proper metadata\n");
    }

    /**
     * Step 2: Read raw data with type conversion and transform to business objects
     */
    public static List<SalesRecord> readAndProcessSalesData() {
        System.out.println("2. Reading and processing sales data...");
        
        try {
            // Read typed data from Excel - works from both filesystem and JAR
            List<Map<String, Object>> rawData = ExcelIO.read("examples/raw_sales_data.xlsx")
                .withMetadata("examples/sales_metadata.xlsx")
                .sheet("Sales Export")
                .asMaps();
                
            // Transform to business objects
            List<SalesRecord> salesRecords = rawData.stream()
                .map(row -> new SalesRecord(
                    (String) row.get("Order ID"),
                    (String) row.get("Product SKU"), 
                    (String) row.get("Product Name"),
                    (String) row.get("Category"),
                    (LocalDate) row.get("Sale Date"),
                    (Integer) row.get("Quantity"),
                    (BigDecimal) row.get("Unit Price"),
                    (BigDecimal) row.get("Total Amount"),
                    (String) row.get("Customer Segment"),
                    (Boolean) row.get("Is Premium")
                ))
                .collect(Collectors.toList());
                
            System.out.println("   ✓ Processed " + salesRecords.size() + " sales records\n");
            return salesRecords;
            
        } catch (Exception e) {
            System.out.println("   Error reading sales data: " + e.getMessage());
            return new ArrayList<>();
        }
    }

    /**
     * Step 3: Generate categorized performance report
     */
    public static void generatePerformanceReport(List<SalesRecord> salesData) {
        System.out.println("3. Generating categorized performance report...");
        
        if (salesData.isEmpty()) {
            System.out.println("   No data to process");
            return;
        }
        
        // Group by category and calculate metrics
        Map<String, List<SalesRecord>> byCategory = salesData.stream()
            .collect(Collectors.groupingBy(SalesRecord::getCategory));
        
        var reportWriter = ExcelIO.writeCategorized("examples/sales_performance_report.xlsx")
            .sheet("Category Performance Analysis")
            .header("Category/Product", "Orders", "Units Sold", "Revenue", "Avg Order Value", "Premium %");
        
        for (String category : byCategory.keySet()) {
            List<SalesRecord> categoryRecords = byCategory.get(category);
            
            // Calculate category totals
            int totalOrders = categoryRecords.size();
            int totalUnits = categoryRecords.stream().mapToInt(SalesRecord::getQuantity).sum();
            BigDecimal totalRevenue = categoryRecords.stream()
                .map(SalesRecord::getTotalAmount)
                .reduce(BigDecimal.ZERO, BigDecimal::add);
            double avgOrderValue = totalRevenue.doubleValue() / totalOrders;
            double premiumPercent = categoryRecords.stream()
                .mapToDouble(r -> r.getIsPremium() ? 1.0 : 0.0)
                .average().orElse(0.0) * 100;
                
            // Add category summary
            reportWriter.category(category, totalOrders, totalUnits, totalRevenue, 
                                String.format("$%.2f", avgOrderValue), String.format("%.1f%%", premiumPercent));
            
            // Group by product within category
            Map<String, List<SalesRecord>> byProduct = categoryRecords.stream()
                .collect(Collectors.groupingBy(SalesRecord::getProductName));
                
            for (String productName : byProduct.keySet()) {
                List<SalesRecord> productRecords = byProduct.get(productName);
                
                int productOrders = productRecords.size();
                int productUnits = productRecords.stream().mapToInt(SalesRecord::getQuantity).sum();
                BigDecimal productRevenue = productRecords.stream()
                    .map(SalesRecord::getTotalAmount)
                    .reduce(BigDecimal.ZERO, BigDecimal::add);
                double productAvgOrder = productRevenue.doubleValue() / productOrders;
                double productPremium = productRecords.stream()
                    .mapToDouble(r -> r.getIsPremium() ? 1.0 : 0.0)
                    .average().orElse(0.0) * 100;
                    
                reportWriter.detail("  " + productName, productOrders, productUnits, productRevenue,
                                  String.format("$%.2f", productAvgOrder), String.format("%.1f%%", productPremium));
            }
        }
        
        reportWriter.save();
        System.out.println("   ✓ Created: examples/sales_performance_report.xlsx\n");
    }

    /**
     * Step 4: Generate detailed transaction report with alternating rows
     */
    public static void generateDetailedTransactionReport(List<SalesRecord> salesData) {
        System.out.println("4. Generating detailed transaction report with alternating rows...");
        
        if (salesData.isEmpty()) {
            System.out.println("   No data to process");
            return;
        }
        
        try (InputStream templateStream = ResourceUtil.getInputStream("src/main/resources/examples/header_details_template.xlsx")) {
            // Sort by date for chronological order
            List<SalesRecord> sortedData = salesData.stream()
                .sorted(Comparator.comparing(SalesRecord::getSaleDate))
                .collect(Collectors.toList());
            
            var writer = ExcelIO.alternatingRows(templateStream)
                .sheet("All Transactions")
                .header("Date", "Order ID", "Product", "Category", "Qty", "Unit Price", "Total", "Customer Type");
                
            for (SalesRecord record : sortedData) {
                writer.row(
                    record.getSaleDate().toString(),
                    record.getOrderId(),
                    record.getProductName(),
                    record.getCategory(),
                    record.getQuantity(),
                    record.getUnitPrice(),
                    record.getTotalAmount(),
                    record.getCustomerSegment() + (record.getIsPremium() ? " (Premium)" : "")
                );
            }
            
            writer.saveAs("examples/detailed_transaction_report.xlsx");
            System.out.println("   ✓ Created: examples/detailed_transaction_report.xlsx with alternating row styles");
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }
        System.out.println();
    }

    /**
     * Step 5: Generate executive summary report
     */
    public static void generateExecutiveSummary(List<SalesRecord> salesData) {
        System.out.println("5. Generating executive summary...");
        
        if (salesData.isEmpty()) {
            System.out.println("   No data to process");
            return;
        }
        
        // Calculate overall metrics
        int totalOrders = salesData.size();
        int totalUnits = salesData.stream().mapToInt(SalesRecord::getQuantity).sum();
        BigDecimal totalRevenue = salesData.stream()
            .map(SalesRecord::getTotalAmount)
            .reduce(BigDecimal.ZERO, BigDecimal::add);
        long premiumOrders = salesData.stream().mapToLong(r -> r.getIsPremium() ? 1 : 0).sum();
        
        // Category breakdown
        Map<String, List<SalesRecord>> byCategory = salesData.stream()
            .collect(Collectors.groupingBy(SalesRecord::getCategory));
        
        // Create summary report
        ExcelIO.write("examples/executive_summary.xlsx")
            .sheet("Executive Summary")
            .header("Metric", "Value", "Notes")
            
            // Key Performance Indicators
            .row("Total Orders", totalOrders, "Orders processed in period")
            .row("Total Units Sold", totalUnits, "Individual items sold")
            .row("Total Revenue", totalRevenue, "Gross sales revenue")
            .row("Average Order Value", String.format("$%.2f", totalRevenue.doubleValue() / totalOrders), "Revenue per order")
            .row("Premium Order Rate", String.format("%.1f%%", (double)premiumOrders / totalOrders * 100), "Percentage of premium customers")
            .row("", "", "")  // Separator
            .row("CATEGORY BREAKDOWN", "", "")
            .save();
            
        // Append category performance
        var summaryWriter = ExcelIO.write("examples/category_summary.xlsx")
            .sheet("Category Breakdown")
            .header("Category", "Orders", "Revenue", "Market Share %", "Avg Order Value");
            
        for (String category : byCategory.keySet()) {
            List<SalesRecord> categoryRecords = byCategory.get(category);
            BigDecimal categoryRevenue = categoryRecords.stream()
                .map(SalesRecord::getTotalAmount)
                .reduce(BigDecimal.ZERO, BigDecimal::add);
            double marketShare = categoryRevenue.doubleValue() / totalRevenue.doubleValue() * 100;
            double avgOrderValue = categoryRevenue.doubleValue() / categoryRecords.size();
            
            summaryWriter.row(category, categoryRecords.size(), categoryRevenue, 
                            String.format("%.1f%%", marketShare), String.format("$%.2f", avgOrderValue));
        }
        
        summaryWriter.save();
        
        System.out.println("   ✓ Created: examples/executive_summary.xlsx");
        System.out.println("   ✓ Created: examples/category_summary.xlsx");
        
        // Demonstrate bulk data processing
        demonstrateBulkDataProcessing(salesData);
        System.out.println();
    }
    
    /**
     * Demonstrate processing large datasets efficiently
     */
    public static void demonstrateBulkDataProcessing(List<SalesRecord> salesData) {
        System.out.println("6. Demonstrating bulk data processing...");
        
        // Convert to maps for bulk Excel operations
        List<Map<String, Object>> salesMaps = salesData.stream()
            .map(record -> {
                Map<String, Object> map = new HashMap<>();
                map.put("Order ID", record.getOrderId());
                map.put("Product", record.getProductName());
                map.put("Category", record.getCategory());
                map.put("Date", record.getSaleDate());
                map.put("Quantity", record.getQuantity());
                map.put("Unit Price", record.getUnitPrice());
                map.put("Total", record.getTotalAmount());
                map.put("Customer", record.getCustomerSegment());
                map.put("Premium", record.getIsPremium() ? "Yes" : "No");
                return map;
            })
            .collect(Collectors.toList());
        
        // Create bulk report using collection
        try (InputStream templateStream = ResourceUtil.getInputStream("src/main/resources/examples/alternating_template.xlsx")) {
            ExcelIO.alternatingRows(templateStream)
                .sheet("Bulk Sales Data", salesMaps)  // Pass entire collection
                .saveAs("examples/bulk_sales_export.xlsx");
                
            System.out.println("   ✓ Created: examples/bulk_sales_export.xlsx (" + salesMaps.size() + " records via bulk processing)");
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }
        
        // Demonstrate filtering and processing subsets
        List<Map<String, Object>> premiumSales = salesMaps.stream()
            .filter(map -> "Yes".equals(map.get("Premium")))
            .collect(Collectors.toList());
            
        try (InputStream templateStream = ResourceUtil.getInputStream("src/main/resources/examples/alternating_template.xlsx")) {
            ExcelIO.alternatingRows(templateStream)
                .sheet("Premium Sales Only", premiumSales)
                .saveAs("examples/premium_sales_report.xlsx");
                
            System.out.println("   ✓ Created: examples/premium_sales_report.xlsx (" + premiumSales.size() + " premium records filtered)");
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }
    }

    /**
     * Business object representing a sales record
     */
    public static class SalesRecord {
        private String orderId;
        private String productSku;
        private String productName;
        private String category;
        private LocalDate saleDate;
        private Integer quantity;
        private BigDecimal unitPrice;
        private BigDecimal totalAmount;
        private String customerSegment;
        private Boolean isPremium;

        public SalesRecord(String orderId, String productSku, String productName, String category,
                          LocalDate saleDate, Integer quantity, BigDecimal unitPrice, BigDecimal totalAmount,
                          String customerSegment, Boolean isPremium) {
            this.orderId = orderId;
            this.productSku = productSku;
            this.productName = productName;
            this.category = category;
            this.saleDate = saleDate;
            this.quantity = quantity;
            this.unitPrice = unitPrice;
            this.totalAmount = totalAmount;
            this.customerSegment = customerSegment;
            this.isPremium = isPremium;
        }

        // Getters
        public String getOrderId() { return orderId; }
        public String getProductSku() { return productSku; }
        public String getProductName() { return productName; }
        public String getCategory() { return category; }
        public LocalDate getSaleDate() { return saleDate; }
        public Integer getQuantity() { return quantity; }
        public BigDecimal getUnitPrice() { return unitPrice; }
        public BigDecimal getTotalAmount() { return totalAmount; }
        public String getCustomerSegment() { return customerSegment; }
        public Boolean getIsPremium() { return isPremium; }
    }
}