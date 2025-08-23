package com.gwill.io.excel.examples;

import com.gwill.io.excel.ExcelIO;
import com.gwill.io.excel.util.ResourceUtil;

import java.io.InputStream;
import java.time.LocalDate;
import java.util.*;

/**
 * Examples demonstrating categorized Excel reporting capabilities.
 * Shows how to create hierarchical reports with summary and detail rows.
 */
public class CategorizedReportExamples {

    public static void main(String[] args) {
        System.out.println("=== Excel-IO Categorized Reporting Examples ===\n");

        // Example 1: Basic categorized report
        createBasicCategorizedReport();

        // Example 2: Financial summary report
        createFinancialSummaryReport();

        // Example 3: Sales performance report
        createSalesPerformanceReport();

        // Example 4: Template-based report (if template exists)
        createTemplateBasedReport();

        System.out.println("\n=== All categorized reporting examples completed successfully! ===");
    }

    /**
     * Example 1: Basic categorized report with simple data
     */
    public static void createBasicCategorizedReport() {
        System.out.println("1. Creating basic categorized report...");

        ExcelIO.writeCategorized("examples/basic_categorized.xlsx")
            .sheet("Product Categories")
            .header("Product", "Q1 Sales", "Q2 Sales", "Total", "Growth %")

            // Electronics category
            .category("Electronics", 150000, 180000, 330000, 20.0)
                .detail("  Laptops", 80000, 95000, 175000, 18.8)
                .detail("  Phones", 50000, 60000, 110000, 20.0)
                .detail("  Tablets", 20000, 25000, 45000, 25.0)

            // Clothing category
            .category("Clothing", 90000, 85000, 175000, -5.6)
                .detail("  Shirts", 40000, 38000, 78000, -5.0)
                .detail("  Pants", 30000, 27000, 57000, -10.0)
                .detail("  Shoes", 20000, 20000, 40000, 0.0)

            // Books category
            .category("Books", 25000, 30000, 55000, 20.0)
                .detail("  Fiction", 15000, 18000, 33000, 20.0)
                .detail("  Non-fiction", 10000, 12000, 22000, 20.0)

            .save();

        System.out.println("   ✓ Created: examples/basic_categorized.xlsx\n");
    }

    /**
     * Example 2: Financial summary with income and expenses
     */
    public static void createFinancialSummaryReport() {
        System.out.println("2. Creating financial summary report...");

        ExcelIO.writeCategorized("examples/financial_summary.xlsx")
            .encoding("UTF-8")
            .sheet("财务摘要报告")  // Chinese sheet name
            .header("项目", "预算", "实际", "差异", "完成率%")

            // Revenue section
            .category("收入", 1200000, 1350000, 150000, 112.5)
                .detail("  产品销售", 1000000, 1100000, 100000, 110.0)
                .detail("  服务收入", 150000, 180000, 30000, 120.0)
                .detail("  投资收益", 50000, 70000, 20000, 140.0)

            // Operating expenses
            .category("运营支出", 800000, 750000, -50000, 93.8)
                .detail("  人员工资", 500000, 485000, -15000, 97.0)
                .detail("  办公租金", 120000, 120000, 0, 100.0)
                .detail("  设备采购", 100000, 80000, -20000, 80.0)
                .detail("  营销推广", 80000, 65000, -15000, 81.3)

            // Administrative expenses
            .category("管理费用", 200000, 180000, -20000, 90.0)
                .detail("  法律咨询", 50000, 40000, -10000, 80.0)
                .detail("  财务审计", 30000, 30000, 0, 100.0)
                .detail("  保险费用", 70000, 65000, -5000, 92.9)
                .detail("  其他管理", 50000, 45000, -5000, 90.0)

            // Net profit calculation
            .category("净利润", 200000, 420000, 220000, 210.0)

            .save();

        System.out.println("   ✓ Created: examples/financial_summary.xlsx with Chinese labels\n");
    }

    /**
     * Example 3: Sales performance report by region and product
     */
    public static void createSalesPerformanceReport() {
        System.out.println("3. Creating sales performance report...");

        ExcelIO.writeCategorized("examples/sales_performance.xlsx")
            .sheet("Regional Sales Performance")
            .header("Region/Product", "Jan", "Feb", "Mar", "Q1 Total", "Target", "Achievement %")

            // North America region
            .category("North America", 450000, 520000, 580000, 1550000, 1500000, 103.3)
                .detail("  USA", 350000, 400000, 450000, 1200000, 1150000, 104.3)
                .detail("    - Product A", 200000, 230000, 260000, 690000, 650000, 106.2)
                .detail("    - Product B", 150000, 170000, 190000, 510000, 500000, 102.0)
                .detail("  Canada", 100000, 120000, 130000, 350000, 350000, 100.0)
                .detail("    - Product A", 60000, 72000, 78000, 210000, 210000, 100.0)
                .detail("    - Product B", 40000, 48000, 52000, 140000, 140000, 100.0)

            // Europe region
            .category("Europe", 380000, 420000, 460000, 1260000, 1200000, 105.0)
                .detail("  Germany", 150000, 170000, 185000, 505000, 480000, 105.2)
                .detail("  France", 120000, 130000, 140000, 390000, 380000, 102.6)
                .detail("  UK", 110000, 120000, 135000, 365000, 340000, 107.4)

            // Asia Pacific region
            .category("Asia Pacific", 320000, 380000, 420000, 1120000, 1000000, 112.0)
                .detail("  Japan", 150000, 180000, 200000, 530000, 450000, 117.8)
                .detail("  Australia", 100000, 120000, 130000, 350000, 320000, 109.4)
                .detail("  Singapore", 70000, 80000, 90000, 240000, 230000, 104.3)

            // Global totals
            .category("GLOBAL TOTAL", 1150000, 1320000, 1460000, 3930000, 3700000, 106.2)

            .save();

        System.out.println("   ✓ Created: examples/sales_performance.xlsx with regional breakdown\n");
    }

    /**
     * Example 4: Template-based categorized report with custom styling
     */
    public static void createTemplateBasedReport() {
        System.out.println("4. Creating template-based categorized report...");
        System.out.println("   ✓ Using template: examples/categorized_template.xlsx");

        try (InputStream templateStream = ResourceUtil.getInputStream("examples/categorized_template.xlsx")) {
            // Create a report using the pre-built template
            ExcelIO.categorizedFromTemplate(templateStream)
                .sheet("Styled Performance Report")
                .header("Department", "Budget", "Actual", "Variance")

                .category("Engineering", 500000, 485000, -15000)
                    .detail("  Backend Team", 200000, 195000, -5000)
                    .detail("  Frontend Team", 150000, 145000, -5000)
                    .detail("  DevOps Team", 150000, 145000, -5000)

                .category("Marketing", 300000, 320000, 20000)
                    .detail("  Digital Marketing", 150000, 160000, 10000)
                    .detail("  Content Creation", 100000, 110000, 10000)
                    .detail("  Events", 50000, 50000, 0)

                .category("Operations", 200000, 190000, -10000)
                    .detail("  HR", 100000, 95000, -5000)
                    .detail("  Finance", 60000, 58000, -2000)
                    .detail("  Admin", 40000, 37000, -3000)

                .saveAs("examples/styled_department_report.xlsx");

            System.out.println("   ✓ Created: examples/styled_department_report.xlsx with template styling");
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }

        // Example 5: Demonstrate bulk data loading for categorized reports
        createBulkCategorizedReport();
        System.out.println();
    }

    /**
     * Example 5: Demonstrate bulk data loading with categorized reports
     */
    public static void createBulkCategorizedReport() {
        System.out.println("\n5. Creating categorized report with bulk data...");

        // Simulate data from database or API
        Map<String, List<RegionData>> regionSales = createRegionalSalesData();

        var reportWriter = ExcelIO.writeCategorized("examples/bulk_regional_sales.xlsx")
            .sheet("Regional Performance")
            .header("Region/Country", "Q1", "Q2", "Q3", "Q4", "Total", "Growth %");

        // Process each region with its countries
        for (String region : regionSales.keySet()) {
            List<RegionData> countries = regionSales.get(region);

            // Calculate region totals
            double regionQ1 = countries.stream().mapToDouble(RegionData::getQ1).sum();
            double regionQ2 = countries.stream().mapToDouble(RegionData::getQ2).sum();
            double regionQ3 = countries.stream().mapToDouble(RegionData::getQ3).sum();
            double regionQ4 = countries.stream().mapToDouble(RegionData::getQ4).sum();
            double regionTotal = regionQ1 + regionQ2 + regionQ3 + regionQ4;
            double regionGrowth = ((regionQ4 - regionQ1) / regionQ1) * 100;

            // Add region category
            reportWriter.category(region,
                String.format("$%.0fK", regionQ1/1000),
                String.format("$%.0fK", regionQ2/1000),
                String.format("$%.0fK", regionQ3/1000),
                String.format("$%.0fK", regionQ4/1000),
                String.format("$%.0fK", regionTotal/1000),
                String.format("%.1f%%", regionGrowth));

            // Add country details using loop
            for (RegionData country : countries) {
                double countryTotal = country.getQ1() + country.getQ2() + country.getQ3() + country.getQ4();
                double countryGrowth = ((country.getQ4() - country.getQ1()) / country.getQ1()) * 100;

                reportWriter.detail("  " + country.getCountryName(),
                    String.format("$%.0fK", country.getQ1()/1000),
                    String.format("$%.0fK", country.getQ2()/1000),
                    String.format("$%.0fK", country.getQ3()/1000),
                    String.format("$%.0fK", country.getQ4()/1000),
                    String.format("$%.0fK", countryTotal/1000),
                    String.format("%.1f%%", countryGrowth));
            }
        }

        reportWriter.save();
        System.out.println("   ✓ Created: examples/bulk_regional_sales.xlsx with structured data");
    }

    /**
     * Create regional sales data - simulates real business data
     */
    private static Map<String, List<RegionData>> createRegionalSalesData() {
        Map<String, List<RegionData>> regionData = new LinkedHashMap<>();
        Random random = new Random(789);

        // North America
        List<RegionData> northAmerica = Arrays.asList(
            new RegionData("United States", 450000, 520000, 580000, 650000),
            new RegionData("Canada", 120000, 135000, 145000, 160000),
            new RegionData("Mexico", 80000, 90000, 95000, 105000)
        );
        regionData.put("North America", northAmerica);

        // Europe
        List<RegionData> europe = Arrays.asList(
            new RegionData("Germany", 200000, 220000, 240000, 260000),
            new RegionData("France", 150000, 165000, 175000, 190000),
            new RegionData("United Kingdom", 180000, 175000, 185000, 195000),
            new RegionData("Spain", 100000, 110000, 115000, 125000)
        );
        regionData.put("Europe", europe);

        // Asia Pacific
        List<RegionData> asiaPacific = Arrays.asList(
            new RegionData("Japan", 250000, 280000, 300000, 320000),
            new RegionData("Australia", 140000, 150000, 160000, 175000),
            new RegionData("Singapore", 90000, 95000, 100000, 110000),
            new RegionData("South Korea", 110000, 125000, 135000, 150000)
        );
        regionData.put("Asia Pacific", asiaPacific);

        return regionData;
    }

    /**
     * Business object for regional sales data
     */
    private static class RegionData {
        private final String countryName;
        private final double q1, q2, q3, q4;

        public RegionData(String countryName, double q1, double q2, double q3, double q4) {
            this.countryName = countryName;
            this.q1 = q1;
            this.q2 = q2;
            this.q3 = q3;
            this.q4 = q4;
        }

        public String getCountryName() { return countryName; }
        public double getQ1() { return q1; }
        public double getQ2() { return q2; }
        public double getQ3() { return q3; }
        public double getQ4() { return q4; }
    }
}
