package com.gwill.io.excel.examples;

/**
 * Utility class to run all Excel-IO examples in sequence.
 * This provides an easy way to see all library capabilities in action.
 */
public class ExampleRunner {

    public static void main(String[] args) {
        System.out.println("===============================================");
        System.out.println("🚀 Running All Excel-IO Examples");
        System.out.println("===============================================\n");

        try {
            // Run all example classes in logical order
            runExample("Basic Usage Examples", BasicUsageExamples::main, args);
            runExample("Typed Reading Examples", TypedReadingExamples::main, args);
            runExample("Categorized Report Examples", CategorizedReportExamples::main, args);
            runExample("Alternating Rows Examples", AlternatingRowsExamples::main, args);
            runExample("Form Template Examples", FormTemplateExamples::main, args);
            runExample("Real-World Example", RealWorldExample::main, args);

            System.out.println("===============================================");
            System.out.println("✅ All examples completed successfully!");
            System.out.println("📁 Check the 'examples/' directory for generated files");
            System.out.println("===============================================");

        } catch (Exception e) {
            System.err.println("❌ Error running examples: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }

    /**
     * Run a single example with error handling and timing
     */
    private static void runExample(String name, ExampleRunnable example, String[] args) {
        System.out.println("🔄 Running: " + name);
        System.out.println("----------------------------------------");

        long startTime = System.currentTimeMillis();

        try {
            example.run(args);
            long duration = System.currentTimeMillis() - startTime;
            System.out.println("✅ " + name + " completed in " + duration + "ms\n");

        } catch (Exception e) {
            System.err.println("❌ " + name + " failed: " + e.getMessage());
            throw new RuntimeException("Example failed: " + name, e);
        }
    }

    /**
     * Functional interface for example main methods
     */
    @FunctionalInterface
    private interface ExampleRunnable {
        void run(String[] args) throws Exception;
    }
}
