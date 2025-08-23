package com.gwill.io.excel.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

/**
 * Utility class for reading files from both filesystem and classpath resources.
 * This allows templates to work both in development (filesystem) and production (JAR).
 */
public class ResourceUtil {
    
    /**
     * Get InputStream for a file, trying filesystem first, then classpath resources.
     * 
     * @param relativePath Path relative to project root (filesystem) or classpath root (resources)
     * @return InputStream for the file
     * @throws IllegalStateException if file not found in either location
     */
    public static InputStream getInputStream(String relativePath) {
        // Try filesystem first (development environment)
        try {
            return getInputStreamFromFileSystem(relativePath);
        } catch (FileNotFoundException e) {
            // Fallback to classpath resources (JAR/production environment)
            return getInputStreamFromClasspath(relativePath);
        }
    }
    
    /**
     * Get InputStream from filesystem
     */
    private static InputStream getInputStreamFromFileSystem(String relativePath) throws FileNotFoundException {
        File file = new File(relativePath);
        if (file.exists()) {
            return new FileInputStream(file);
        }
        throw new FileNotFoundException("File not found in filesystem: " + relativePath);
    }
    
    /**
     * Get InputStream from classpath resources
     */
    private static InputStream getInputStreamFromClasspath(String relativePath) {
        // Try with ClassLoader (no leading slash)
        InputStream is = ResourceUtil.class.getClassLoader().getResourceAsStream(relativePath);
        if (is != null) {
            return is;
        }
        
        // Try with Class (with leading slash)
        is = ResourceUtil.class.getResourceAsStream("/" + relativePath);
        if (is != null) {
            return is;
        }
        
        throw new IllegalStateException("Resource not found in classpath: " + relativePath);
    }
    
    /**
     * Check if a resource exists in either filesystem or classpath
     */
    public static boolean resourceExists(String relativePath) {
        try {
            getInputStream(relativePath).close();
            return true;
        } catch (Exception e) {
            return false;
        }
    }
}