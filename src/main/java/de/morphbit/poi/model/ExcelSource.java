package de.morphbit.poi.model;

import de.morphbit.poi.exception.ExcelSourceNotSupportedException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

public abstract class ExcelSource {

    private String password;
    private boolean readonly;

    public ExcelSource() {
        this( null, false);
    }

    public ExcelSource(String password) {
        this( password, false);
    }

    public ExcelSource(boolean readonly) {
        this( null, readonly);
    }

    public ExcelSource(String password, boolean readonly) {
        this.password = password;
        this.readonly = readonly;
    }

    public abstract Object getSource();

    public String getPassword() {
        return password;
    }

    public void setPassword(String password) {
        this.password = password;
    }

    public boolean isReadonly() {
        return readonly;
    }

    public void setReadonly(boolean readonly) {
        this.readonly = readonly;
    }

    public Workbook openWorkbook() throws IOException, InvalidFormatException, ExcelSourceNotSupportedException {
        Object source = getSource();
        if(source == null) {
            throw  new NullPointerException("Field 'source' must not be null");
        } else if(InputStream.class.equals(source.getClass())) {
            return WorkbookFactory.create((InputStream)source);
        } else if (File.class.equals(source.getClass())) {
            return WorkbookFactory.create((File)source);
        } else {
            throw new ExcelSourceNotSupportedException("ExcelSource of type " + source.getClass().getName() + " not supported");
        }
    }
}
