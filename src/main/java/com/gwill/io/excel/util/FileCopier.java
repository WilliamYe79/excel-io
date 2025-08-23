package com.gwill.io.excel.util;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * File copying utility class for handling POI workbook style conflicts.
 * 
 * <p>This class solves the classic Apache POI problem: "This Style does not belong to the 
 * supplied Workbook Styles Source. Are you trying to assign a style from one workbook 
 * to the cell of a different workbook?"</p>
 * 
 * <p>By using in-memory byte array operations, all workbook manipulations happen within 
 * the same workbook object, completely avoiding style conflicts.</p>
 * 
 * @author William YE of G-WILL Team
 */
public class FileCopier {
    
    /**
     * Copy a file to memory as a byte array.
     * 
     * @param filePath the path to the file to copy
     * @return the file content as a byte array
     * @throws IOException if an error occurs while reading the file
     */
    public static byte[] copyFileToMemoryAsByteArray(String filePath) throws IOException {
        if (filePath == null || filePath.trim().isEmpty()) {
            throw new IllegalArgumentException("File path cannot be null or empty");
        }
        return Files.readAllBytes(Path.of(filePath));
    }
    
    /**
     * Copy an InputStream's content to memory as a byte array.
     * This is the key method for avoiding POI style conflicts.
     * 
     * @param inputStream the input stream to copy
     * @return the stream content as a byte array
     * @throws IOException if an error occurs while reading the stream
     */
    public static byte[] copyStreamToMemoryAsByteArray(InputStream inputStream) throws IOException {
        if (inputStream == null) {
            throw new IllegalArgumentException("InputStream cannot be null");
        }
        
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            byte[] buffer = new byte[8192];
            int bytesRead;
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                baos.write(buffer, 0, bytesRead);
            }
            return baos.toByteArray();
        }
    }
    
    /**
     * Write a byte array to a file.
     * Creates parent directories if they don't exist.
     * 
     * @param data the byte array to write
     * @param filePath the target file path
     * @throws IOException if an error occurs while writing the file
     */
    public static void writeByteArrayToFile(byte[] data, String filePath) throws IOException {
        if (data == null) {
            throw new IllegalArgumentException("Data cannot be null");
        }
        if (filePath == null || filePath.trim().isEmpty()) {
            throw new IllegalArgumentException("File path cannot be null or empty");
        }
        
        Path path = Path.of(filePath);
        // Create parent directories if they don't exist
        Path parentDir = path.getParent();
        if (parentDir != null) {
            Files.createDirectories(parentDir);
        }
        
        Files.write(path, data);
    }
    
    /**
     * Create an InputStream from a byte array.
     * Used to create new input streams from in-memory template data.
     * 
     * @param data the byte array data
     * @return an InputStream based on the byte array
     */
    public static InputStream createInputStreamFromByteArray(byte[] data) {
        if (data == null) {
            throw new IllegalArgumentException("Data cannot be null");
        }
        return new ByteArrayInputStream(data);
    }
}