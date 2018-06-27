package de.morphbit.poi;

import de.morphbit.poi.exception.ExcelReadException;
import de.morphbit.poi.exception.ExcelSourceNotSupportedException;
import de.morphbit.poi.mapper.ExcelRowMapper;
import de.morphbit.poi.mapper.ExcelRowMapperWithHeader;
import de.morphbit.poi.model.ExcelSource;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelDataTemplate {

    public static final ExcelRowMapper<Map<String, Integer>> DEFAULT_HEADER_MAPPER =
            row -> {
                Map<String, Integer> headers = new HashMap<>();
                row.forEach(cell -> {
                    headers.put(cell.getRichStringCellValue().getString(),
                            cell.getColumnIndex());
                });
                return headers;
            };

    private ExcelSource source;
    private ExcelRowMapper<Map<String, Integer>> headerMapper;


    public ExcelDataTemplate() {
        this(null, null);
    }

    public ExcelDataTemplate(ExcelSource source) {
        this(source, null);
    }

    public ExcelDataTemplate(ExcelRowMapper<Map<String, Integer>> headerMapper) {
        this(null, headerMapper);
    }

    public ExcelDataTemplate(ExcelSource source, ExcelRowMapper<Map<String, Integer>> headerMapper) {
        this.source = source;
        this.headerMapper = headerMapper;
    }

    public <T> List<T> read(int sheetIndex,
                            ExcelRowMapper<T> rowMapper, boolean ignoreFirstRow)
            throws ExcelReadException, IOException {
        return read(this.source, sheetIndex, rowMapper, ignoreFirstRow);
    }

    public <T> List<T> read(ExcelSource source, int sheetIndex,
                            ExcelRowMapper<T> rowMapper, boolean ignoreFirstRow)
            throws ExcelReadException, IOException {
        try (Workbook workbook = openWorkbook(source)) {
            return doReadListFromSheet(getSheet(workbook, sheetIndex), rowMapper, ignoreFirstRow);
        }
    }

    public <T> List<T> read(String sheetName,
                            ExcelRowMapper<T> rowMapper, boolean ignoreFirstRow)
            throws ExcelReadException, IOException {
        return read(this.source, sheetName, rowMapper, ignoreFirstRow);
    }

    public <T> List<T> read(ExcelSource source, String sheetName,
                            ExcelRowMapper<T> rowMapper, boolean ignoreFirstRow)
            throws ExcelReadException, IOException {
        try (Workbook workbook = openWorkbook(source)) {
            return doReadListFromSheet(getSheet(workbook, sheetName), rowMapper, ignoreFirstRow);
        }
    }

    public <T> List<T> read(int sheetIndex,
                            ExcelRowMapperWithHeader<T> rowMapper)
            throws ExcelReadException, IOException {
        return read(this.source, sheetIndex, rowMapper);
    }

    public <T> List<T> read(ExcelSource source, int sheetIndex,
                            ExcelRowMapperWithHeader<T> rowMapper)
            throws ExcelReadException, IOException {
        try (Workbook workbook = openWorkbook(source)) {
            return doReadListFromSheet(getSheet(workbook, sheetIndex), rowMapper);
        }
    }

    public <T> List<T> read(String sheetName,
                            ExcelRowMapperWithHeader<T> rowMapper)
            throws ExcelReadException, IOException {
        return read(this.source, sheetName, rowMapper);
    }

    public <T> List<T> read(ExcelSource source, String sheetName,
                            ExcelRowMapperWithHeader<T> rowMapper) throws ExcelReadException, IOException {
        try (Workbook workbook = openWorkbook(source)) {
            return doReadListFromSheet(getSheet(workbook, sheetName), rowMapper);
        }
    }

    private Workbook openWorkbook(ExcelSource source) throws ExcelReadException {
        if (source == null) {
            throw new NullPointerException("ExcelSource is not set");
        }
        try {
            return source.openWorkbook();
        } catch (InvalidFormatException | ExcelSourceNotSupportedException | IOException e) {
            throw new ExcelReadException("Could not open workbook", e);
        }
    }

    private <T> List<T> doReadListFromSheet(Sheet sheet, ExcelRowMapper<T> rowMapper, boolean ignoreFirstRow) {
        List<T> objects = new ArrayList<>();
        for (Row row : sheet) {
            if (row.getRowNum() > 0 || !ignoreFirstRow) {
                objects.add(rowMapper.map(row));
            }
        }
        return objects;
    }

    private <T> List<T> doReadListFromSheet(Sheet sheet, ExcelRowMapperWithHeader<T> rowMapper) {
        List<T> objects = new ArrayList<>();
        Map<String, Integer> headers = null;
        for (Row row : sheet) {
            if (row.getRowNum() < 1) {
                headers = mapHeaders(row);
            } else {
                objects.add(rowMapper.map(row, headers));
            }
        }
        return objects;
    }

    public Sheet getSheet(Workbook workbook, int index) {
        if (workbook == null) {
            throw new NullPointerException("Parameter 'workbook' must not be null");
        }
        return workbook.getSheetAt(index);
    }

    public Sheet getSheet(Workbook workbook, String name) {
        if (workbook == null) {
            throw new NullPointerException("Parameter 'workbook' must not be null");
        }
        return workbook.getSheet(name);
    }

    private Map<String, Integer> mapHeaders(Row row) {
        ExcelRowMapper<Map<String, Integer>> mapper = DEFAULT_HEADER_MAPPER;
        if (headerMapper != null) {
            mapper = headerMapper;
        }
        return mapper.map(row);
    }
}
