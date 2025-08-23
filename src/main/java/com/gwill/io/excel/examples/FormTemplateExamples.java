package com.gwill.io.excel.examples;

import com.gwill.io.excel.ExcelIO;
import com.gwill.io.excel.util.ResourceUtil;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.time.LocalDate;

/**
 * Examples demonstrating form-based Excel document generation.
 * Shows how to create professional business documents like invoices, purchase orders, and reports.
 */
public class FormTemplateExamples {

    public static void main(String[] args) {
        System.out.println("=== Excel-IO Form Template Examples ===\n");

        // Example 1: Create an invoice using form template
        createInvoiceExample();

        // Example 2: Create a purchase order
        createPurchaseOrderExample();

        // Example 3: Create a service report
        createServiceReportExample();

        // Example 4: Backend integration example (stream output)
        createBackendIntegrationExample();

        System.out.println("\n=== All form template examples completed successfully! ===");
    }

    /**
     * Example 1: Create a professional invoice using FormTemplateWriter
     */
    public static void createInvoiceExample() {
        System.out.println("1. Creating professional invoice...");
        System.out.println("   ✓ Using template: examples/form_template.xlsx");

        try (InputStream templateStream = ResourceUtil.getInputStream("examples/form_template.xlsx")) {
            ExcelIO.formTemplate(templateStream)
                // Form header information
                .setValue("companyName", "Acme Technology Solutions")
                .setValue("companyAddress", "123 Innovation Drive, Tech City, CA 94025")
                .setValue("companyPhone", "+1 (555) 123-4567")
                .setValue("companyEmail", "billing@acmetech.com")

                // Invoice specific details
                .setValue("invoiceNumber", "INV-2024-001")
                .setValue("invoiceDate", LocalDate.now().toString())
                .setValue("dueDate", LocalDate.now().plusDays(30).toString())
//                .setValue("reference", "Project Alpha Development")

                // Customer information
                .setValue("customerName", "Global Manufacturing Corp")
                .setValue("customerContact", "Jane Smith, CTO")
                .setValue("customerEmail", "jane.smith@globalmanufacturing.com")
//                .setValue("projectName", "Enterprise Software Implementation")
                .setValue("poNumber", "PO-2024-GMC-789")
//                .setValue("salesRep", "Robert Chen")
                .setValue("taxId", "12-3456789")

                // Line items for services/products
                .lineItem("DEV-001", "Backend Development (40 hours)", 40, 150.00, 6000.00)
                .lineItem("DEV-002", "Frontend Development (32 hours)", 32, 140.00, 4480.00)
                .lineItem("DEV-003", "Database Design & Setup", 1, 2500.00, 2500.00)
                .lineItem("DEV-004", "Testing & Quality Assurance", 1, 1500.00, 1500.00)
                .lineItem("DEV-005", "Project Management", 1, 3000.00, 3000.00)

                // Financial calculations
                .setValue("subtotal", "17,480.00")
                .setValue("tax", "1,748.00")
                .setValue("grandTotal", new BigDecimal("19228.00") )
                .setValue("paymentTerms", "Net 30 days")
                .setValue("deliveryTime", "The whole project will be finished within 1 month assuming all payments are made in time.")
                .setValue("notes", "Thank you for choosing Acme Technology Solutions!")

                .saveAs("examples/generated_invoice.xlsx");

            System.out.println("   ✓ Created: examples/generated_invoice.xlsx with professional formatting\n");
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }
    }

    /**
     * Example 2: Create a purchase order
     */
    public static void createPurchaseOrderExample() {
        System.out.println("2. Creating purchase order...");

        try (InputStream templateStream = ResourceUtil.getInputStream("examples/form_template.xlsx")) {
            ExcelIO.formTemplate(templateStream)
                // Company header
                .setValue("companyName", "Global Manufacturing Corp")
                .setValue("companyAddress", "456 Industrial Blvd, Manufacturing City, TX 77001")
                .setValue("companyPhone", "+1 (555) 987-6543")
                .setValue("companyEmail", "purchasing@globalmanufacturing.com")

                // Purchase order details
                .setValue("invoiceNumber", "PO-2024-002")  // Reusing field for PO number
                .setValue("invoiceDate", LocalDate.now().toString())
                .setValue("dueDate", LocalDate.now().plusDays(14).toString())  // Delivery date
                .setValue("reference", "Q4 Office Supplies")

                // Vendor information
                .setValue("customerName", "Office Depot Business Solutions")
                .setValue("customerContact", "Mike Rodriguez, Account Manager")
//                .setValue("customerEmail", "mike.rodriguez@officedepot.com")
//                .setValue("projectName", "Office Equipment Procurement")
                .setValue("poNumber", "REQ-2024-456")  // Internal requisition
//                .setValue("salesRep", "Lisa Thompson")
                .setValue("taxId", "98-7654321")

                // Purchase items
                .lineItem("PEN-001", "Blue Ballpoint Pens (Pack of 12)", 50, 8.99, 449.50)
                .lineItem("PAP-001", "A4 Copy Paper (500 sheets)", 25, 12.99, 324.75)
                .lineItem("BIN-001", "3-Ring Binders (2 inch)", 30, 15.99, 479.70)
                .lineItem("STA-001", "Stapler Heavy Duty", 5, 45.99, 229.95)
                .lineItem("FOL-001", "File Folders (Pack of 100)", 10, 18.99, 189.90)

                // Purchase totals
                .setValue("subtotal", "1,673.80")
                .setValue("tax", "167.38")
                .setValue("grandTotal", "1,841.18")
                .setValue("paymentTerms", "Net 15 days")
                .setValue( "deliveryTime", "The whole project will be finished within 1 month assuming all payments are made in time." )
                .setValue("notes", "Please deliver to main warehouse dock.")

                .saveAs("examples/generated_purchase_order.xlsx");

            System.out.println("   ✓ Created: examples/generated_purchase_order.xlsx\n");
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }
    }

    /**
     * Example 3: Create a service report
     */
    public static void createServiceReportExample() {
        System.out.println("3. Creating service report...");

        try (InputStream templateStream = ResourceUtil.getInputStream("examples/form_template.xlsx")) {
            ExcelIO.formTemplate(templateStream)
                // Service company header
                .setValue("companyName", "TechSupport Pro Services")
                .setValue("companyAddress", "789 Support Lane, Service City, NY 10001")
                .setValue("companyPhone", "+1 (555) 444-TECH")
                .setValue("companyEmail", "reports@techsupportpro.com")

                // Report details
                .setValue("invoiceNumber", "RPT-2024-15")
                .setValue("invoiceDate", LocalDate.now().toString())
                .setValue("dueDate", "N/A")  // Not applicable for reports
                .setValue("reference", "Monthly Maintenance Report")

                // Client information
                .setValue("customerName", "Downtown Medical Center")
                .setValue("customerContact", "Dr. Sarah Wilson, IT Director")
                .setValue("customerEmail", "s.wilson@downtownmedical.com")
                .setValue("projectName", "IT Infrastructure Maintenance")
                .setValue("poNumber", "CONTRACT-2024-ITM")
                .setValue("salesRep", "Alex Kim")
                .setValue("taxId", "55-9876543")

                // Service activities performed
                .lineItem("SRV-001", "Server Maintenance & Updates", 8, 125.00, 1000.00)
                .lineItem("SRV-002", "Network Security Audit", 4, 200.00, 800.00)
                .lineItem("SRV-003", "Backup System Verification", 2, 150.00, 300.00)
                .lineItem("SRV-004", "Workstation Health Checks", 12, 75.00, 900.00)
                .lineItem("SRV-005", "Emergency Support Calls", 3, 250.00, 750.00)

                // Service summary
                .setValue("subtotal", new BigDecimal("3750.00") )
                .setValue("tax", "N/A")  // Service contract, no additional tax
                .setValue("grandTotal", new BigDecimal("3750.00") )
                .setValue("paymentTerms", "Monthly Contract")
                .setValue( "deliveryTime", "All services will be finished within 1 month assuming all payments are made in time." )
                .setValue("notes", "All systems operating within normal parameters. Next maintenance scheduled for next month.")

                .saveAs("examples/generated_service_report.xlsx");

            System.out.println("   ✓ Created: examples/generated_service_report.xlsx\n");
        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }
    }

    /**
     * Example 4: Backend integration - demonstrate stream output for web applications
     */
    public static void createBackendIntegrationExample() {
        System.out.println("4. Demonstrating backend integration with stream output...");

        try (InputStream templateStream = ResourceUtil.getInputStream("examples/form_template.xlsx")) {
            // Example 1: Generate to byte array (useful for APIs)
            byte[] invoiceBytes = ExcelIO.formTemplate(templateStream)
                .setValue("companyName", "API Solutions Inc")
                .setValue("invoiceNumber", "API-INV-001")
                .setValue("invoiceDate", LocalDate.now().toString())
                .setValue("customerName", "Web Client Corp")
                .lineItem("API-001", "API Development", 40, 100.00, 4000.00)
                .lineItem("API-002", "Documentation", 8, 75.00, 600.00)
                .setValue("subtotal", "4,600.00")
                .setValue("tax", "460.00")
                .setValue("grandTotal", "5,060.00")
                .toByteArray();

            System.out.println("   ✓ Generated invoice as byte array: " + invoiceBytes.length + " bytes");

            // Example 2: Generate to output stream (useful for HTTP responses)
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            try (InputStream templateStream2 = ResourceUtil.getInputStream("examples/form_template.xlsx")) {
                ExcelIO.formTemplate(templateStream2)
                    .setValue("companyName", "Stream Processing Corp")
                    .setValue("invoiceNumber", "STREAM-001")
                    .setValue("invoiceDate", LocalDate.now().toString())
                    .setValue("customerName", "HTTP Client Ltd")
                    .lineItem("WEB-001", "Web Service Development", 20, 150.00, 3000.00)
                    .setValue("subtotal", "3,000.00")
                    .setValue("tax", "300.00")
                    .setValue("grandTotal", "3,300.00")
                    .writeTo(outputStream);
            }

            System.out.println("   ✓ Generated to output stream: " + outputStream.size() + " bytes");
            System.out.println("   ✓ Perfect for web applications and HTTP responses\n");

        } catch (Exception e) {
            System.out.println("   Error: " + e.getMessage());
        }
    }
}
