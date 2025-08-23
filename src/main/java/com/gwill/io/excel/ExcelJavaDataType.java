package com.gwill.io.excel;

/**
 * Enum representing supported Java data types for Excel I/O operations.
 * This provides a bridge between Java class names in metadata files and internal type processing.
 * 
 * <p>This enum is primarily for internal use within the excel-io library.
 * External users should use standard Java class names (e.g., "java.lang.String") in metadata files.</p>
 */
public enum ExcelJavaDataType {
    STRING("java.lang.String"),
    BOOLEAN("java.lang.Boolean"),
    INTEGER("java.lang.Integer"),
    LONG("java.lang.Long"),
    SHORT("java.lang.Short"),
    BYTE("java.lang.Byte"),
    FLOAT("java.lang.Float"),
    DOUBLE("java.lang.Double"),
    BIG_DECIMAL("java.math.BigDecimal"),
    LOCAL_DATE("java.time.LocalDate"),
    LOCAL_TIME("java.time.LocalTime"),
    LOCAL_DATE_TIME("java.time.LocalDateTime");

    private final String javaClassName;

    ExcelJavaDataType(String javaClassName) {
        this.javaClassName = javaClassName;
    }

    /**
     * Get the Java class name associated with this enum.
     * 
     * @return the Java class name (e.g., "java.lang.String")
     */
    public String getJavaClassName() {
        return javaClassName;
    }

    @Override
    public String toString() {
        return javaClassName;
    }

    /**
     * Find the ExcelJavaDataType that corresponds to the given Java class name.
     * This is the bridge method that converts from metadata file values to internal enum.
     * 
     * @param javaClassName the Java class name from metadata (e.g., "java.lang.String")
     * @return the corresponding ExcelJavaDataType enum
     * @throws IllegalArgumentException if the class name is not supported
     */
    public static ExcelJavaDataType fromClassName(String javaClassName) {
        if (javaClassName == null) {
            throw new IllegalArgumentException("Java class name cannot be null");
        }

        // Normalize the class name
        String normalizedClassName = javaClassName.trim();

        // Find matching enum
        for (ExcelJavaDataType dataType : ExcelJavaDataType.values()) {
            if (dataType.javaClassName.equals(normalizedClassName)) {
                return dataType;
            }
        }

        // If not found, provide helpful error message
        StringBuilder supportedTypes = new StringBuilder();
        for (ExcelJavaDataType dataType : ExcelJavaDataType.values()) {
            if (supportedTypes.length() > 0) {
                supportedTypes.append(", ");
            }
            supportedTypes.append(dataType.javaClassName);
        }

        throw new IllegalArgumentException(
            "Unsupported Java class: '" + normalizedClassName + "'. " +
            "Supported types are: " + supportedTypes.toString()
        );
    }

    /**
     * Get the actual Java Class object for this data type.
     * 
     * @return the Class object
     * @throws ClassNotFoundException if the class cannot be loaded (shouldn't happen for supported types)
     */
    public Class<?> getJavaClass() throws ClassNotFoundException {
        return Class.forName(javaClassName);
    }
}