package de.morphbit.poi.model;

import java.io.File;

public class ExcelFileSource extends ExcelSource {

    private final File source;

    public ExcelFileSource(File source) {
        this.source = source;
    }

    public ExcelFileSource(File source, String password) {
        super(password);
        this.source = source;
    }

    public ExcelFileSource(File source, boolean readonly) {
        super(readonly);
        this.source = source;
    }

    public ExcelFileSource(File source, String password, boolean readonly) {
        super(password, readonly);
        this.source = source;
    }

    @Override
    public Object getSource() {
        return source;
    }
}
