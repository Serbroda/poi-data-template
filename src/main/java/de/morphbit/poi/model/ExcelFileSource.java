package de.morphbit.poi.model;

import java.io.File;

public class ExcelFileSource extends ExcelSource {

    private final File source;

    public ExcelFileSource(File source) {
        this.source = source;
    }

    @Override
    public Object getSource() {
        return source;
    }
}
