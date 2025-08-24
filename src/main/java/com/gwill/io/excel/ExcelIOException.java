package com.gwill.io.excel;

/**
 * Runtime exception thrown when Excel I/O operations fail.
 * This is an unchecked exception to keep the API clean and simple.
 *
 * @author William Ye
 * @version 1.0.0
 */
public class ExcelIOException extends RuntimeException {

    public ExcelIOException(String message) {
        super(message);
    }

    public ExcelIOException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelIOException(Throwable cause) {
        super(cause);
    }
}
