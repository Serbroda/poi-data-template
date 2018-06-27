package de.morphbit.poi.exception;

public class ExcelSourceNotSupportedException extends Exception {

    public ExcelSourceNotSupportedException() {
    }

    public ExcelSourceNotSupportedException(String message) {
        super(message);
    }

    public ExcelSourceNotSupportedException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelSourceNotSupportedException(Throwable cause) {
        super(cause);
    }

    public ExcelSourceNotSupportedException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
