package de.morphbit.poi.exception;

public class ExcelReadException extends Exception {

    public ExcelReadException() {
    }

    public ExcelReadException(String message) {
        super(message);
    }

    public ExcelReadException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelReadException(Throwable cause) {
        super(cause);
    }

    public ExcelReadException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
