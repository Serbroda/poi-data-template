package de.morphbit.poi;

import de.morphbit.poi.exception.ExcelReadException;
import de.morphbit.poi.exception.ExcelSourceNotSupportedException;
import de.morphbit.poi.mapper.ExcelRowMapper;
import de.morphbit.poi.mapper.ExcelRowMapperWithHeader;
import de.morphbit.poi.model.ExcelDataTemplateOptions;
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
import java.util.function.Predicate;
import java.util.stream.Collectors;

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
            ExcelRowMapper<T> rowMapper)
            		throws ExcelReadException, IOException {
    	return read(this.source, sheetIndex, rowMapper, null);
    }

    public <T> List<T> read(ExcelSource source, int sheetIndex,
            ExcelRowMapper<T> rowMapper)
            		throws ExcelReadException, IOException {
    	try (Workbook workbook = openWorkbook(source)) {
    		return doReadListFromSheet(getSheet(workbook, sheetIndex), rowMapper, null);
    	}
    }

    public <T> List<T> read(int sheetIndex,
                            ExcelRowMapper<T> rowMapper, ExcelDataTemplateOptions options)
            throws ExcelReadException, IOException {
        return read(this.source, sheetIndex, rowMapper, options);
    }

    public <T> List<T> read(ExcelSource source, int sheetIndex,
                            ExcelRowMapper<T> rowMapper, ExcelDataTemplateOptions options)
            throws ExcelReadException, IOException {
        try (Workbook workbook = openWorkbook(source)) {
            return doReadListFromSheet(getSheet(workbook, sheetIndex), rowMapper, options);
        }
    }
    
    public <T> List<T> read(String sheetName,
            ExcelRowMapper<T> rowMapper)
            		throws ExcelReadException, IOException {
    	return read(this.source, sheetName, rowMapper, null);
    }

    public <T> List<T> read(ExcelSource source, String sheetName,
            ExcelRowMapper<T> rowMapper)
            		throws ExcelReadException, IOException {
    	try (Workbook workbook = openWorkbook(source)) {
    		return doReadListFromSheet(getSheet(workbook, sheetName), rowMapper, null);
    	}
    }

    public <T> List<T> read(String sheetName,
                            ExcelRowMapper<T> rowMapper, ExcelDataTemplateOptions options)
            throws ExcelReadException, IOException {
        return read(this.source, sheetName, rowMapper, options);
    }

    public <T> List<T> read(ExcelSource source, String sheetName,
                            ExcelRowMapper<T> rowMapper, ExcelDataTemplateOptions options)
            throws ExcelReadException, IOException {
        try (Workbook workbook = openWorkbook(source)) {
            return doReadListFromSheet(getSheet(workbook, sheetName), rowMapper, options);
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
        return read(source, sheetIndex, rowMapper, null);
    }
    
    public <T> List<T> read(int sheetIndex,
            ExcelRowMapperWithHeader<T> rowMapper, ExcelDataTemplateOptions options)
            		throws ExcelReadException, IOException {
    	return read(this.source, sheetIndex, rowMapper);
    }

    public <T> List<T> read(ExcelSource source, int sheetIndex,
            ExcelRowMapperWithHeader<T> rowMapper, ExcelDataTemplateOptions options)
            		throws ExcelReadException, IOException {
    	try (Workbook workbook = openWorkbook(source)) {
    		return doReadListFromSheet(getSheet(workbook, sheetIndex), rowMapper, options);
    	}
    }

    public <T> List<T> read(String sheetName,
                            ExcelRowMapperWithHeader<T> rowMapper)
            throws ExcelReadException, IOException {
        return read(this.source, sheetName, rowMapper);
    }

    public <T> List<T> read(ExcelSource source, String sheetName,
                            ExcelRowMapperWithHeader<T> rowMapper) throws ExcelReadException, IOException {
        return read(source, sheetName, rowMapper, null);
    }
    
    public <T> List<T> read(String sheetName,
            ExcelRowMapperWithHeader<T> rowMapper, ExcelDataTemplateOptions options)
            		throws ExcelReadException, IOException {
    	return read(this.source, sheetName, rowMapper, options);
    }

    public <T> List<T> read(ExcelSource source, String sheetName,
            ExcelRowMapperWithHeader<T> rowMapper, ExcelDataTemplateOptions options) throws ExcelReadException, IOException {
    	try (Workbook workbook = openWorkbook(source)) {
    		return doReadListFromSheet(getSheet(workbook, sheetName), rowMapper, options);
    	}
    }
    
    public <T> Map<String, Integer> readFirstLineAsHeader(String sheetName) throws ExcelReadException, IOException {
    	return readFirstLineAsHeader(sheetName, DEFAULT_HEADER_MAPPER);
    }
    
    public <T> Map<String, Integer> readFirstLineAsHeader(String sheetName, ExcelRowMapper<Map<String, Integer>> headerMapper)
    		throws ExcelReadException, IOException {
    	return readFirstLineAsHeader(this.source, sheetName, headerMapper);
    }
    
    public <T> Map<String, Integer> readFirstLineAsHeader(ExcelSource source, String sheetName) throws ExcelReadException, IOException {
		return readFirstLineAsHeader(source, sheetName, DEFAULT_HEADER_MAPPER);
	}

    public <T> Map<String, Integer> readFirstLineAsHeader(ExcelSource source, String sheetName, ExcelRowMapper<Map<String, Integer>> headerMapper) throws ExcelReadException, IOException {
    	try (Workbook workbook = openWorkbook(source)) {
    		Sheet sheet = getSheet(workbook, sheetName);
    		Row row = sheet.getRow(0);
    		return headerMapper.map(row);
    	}
    }
    
    public <T> Map<String, Integer> readFirstLineAsHeader(int sheetIndex) throws ExcelReadException, IOException {
    	return readFirstLineAsHeader(sheetIndex, DEFAULT_HEADER_MAPPER);
    }
    
    public <T> Map<String, Integer> readFirstLineAsHeader(int sheetIndex, ExcelRowMapper<Map<String, Integer>> headerMapper)
    		throws ExcelReadException, IOException {
    	return readFirstLineAsHeader(this.source, sheetIndex, headerMapper);
    }
    
    public <T> Map<String, Integer> readFirstLineAsHeader(ExcelSource source, int sheetIndex) throws ExcelReadException, IOException {
		return readFirstLineAsHeader(source, sheetIndex, DEFAULT_HEADER_MAPPER);
	}

    public <T> Map<String, Integer> readFirstLineAsHeader(ExcelSource source, int sheetIndex, ExcelRowMapper<Map<String, Integer>> headerMapper) throws ExcelReadException, IOException {
    	try (Workbook workbook = openWorkbook(source)) {
    		Sheet sheet = getSheet(workbook, sheetIndex);
    		Row row = sheet.getRow(0);
    		return headerMapper.map(row);
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

    private <T> List<T> doReadListFromSheet(Sheet sheet, ExcelRowMapper<T> rowMapper, ExcelDataTemplateOptions options) {
        List<T> objects = new ArrayList<>();
        for (Row row : sheet) {
        	if(options == null || row.getRowNum() + 1 > options.getIgnoreFirstLinesCount()) {
        		objects.add(rowMapper.map(row));
        	}
        }
        return objects;
    }

    @SuppressWarnings("unchecked")
	private <T> List<T> doReadListFromSheet(Sheet sheet, ExcelRowMapperWithHeader<T> rowMapper, ExcelDataTemplateOptions options) {
        List<T> objects = new ArrayList<>();
        Map<String, Integer> headers = null;
        boolean headerRead = false;
        for (Row row : sheet) {
        	if(options == null || row.getRowNum() + 1 > options.getIgnoreFirstLinesCount()) {
        		if (!headerRead) {
                    headers = mapHeaders(row);
                    headerRead = true;
                } else {
                	T obj = rowMapper.map(row, headers);
                	if(options.getFilter() == null || (options.getFilter() != null && ((Predicate<? super T>)options.getFilter()).test(obj))) {
                		objects.add(rowMapper.map(row, headers));
                	}                
                }
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
