package de.morphbit.poi.model;

import java.io.InputStream;

public class ExcelInputStreamSource extends ExcelSource {

    private final InputStream source;

    public ExcelInputStreamSource(InputStream source) {
        this.source = source;
    }

    @Override
    public Object getSource() {
        return source;
    }
}
