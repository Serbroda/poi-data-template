package de.morphbit.poi;

import de.morphbit.poi.mapper.ExcelRowMapper;
import de.morphbit.poi.mapper.ExcelRowMapperWithHeader;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
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

    private final ExcelRowMapper<Map<String, Integer>> headerMapper;

    public ExcelDataTemplate() {
        this(null);
    }

    public ExcelDataTemplate(ExcelRowMapper<Map<String, Integer>> headerMapper) {
        this.headerMapper = headerMapper;
    }


    @SuppressWarnings("resource")
    public <T> List<T> read(File file, int sheetIndex,
                            ExcelRowMapper<T> rowMapper, boolean ignoreFirstRow)
            throws EncryptedDocumentException, InvalidFormatException,
            IOException {
        Workbook workbook = null;
        try {
            workbook = openWorkbook(file);
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            List<T> objects = new ArrayList<>();
            for (Row row : sheet) {
                if (row.getRowNum() > 0 || !ignoreFirstRow) {
                    objects.add(rowMapper.map(row));
                }
            }
            return objects;
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    }

    @SuppressWarnings("resource")
    public <T> List<T> read(InputStream inputStream, int sheetIndex,
                            ExcelRowMapper<T> rowMapper, boolean ignoreFirstRow)
            throws Exception {
        Workbook workbook = null;
        try {
            workbook = openWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            List<T> objects = new ArrayList<>();
            for (Row row : sheet) {
                if (row.getRowNum() > 0 || !ignoreFirstRow) {
                    objects.add(rowMapper.map(row));
                }
            }
            return objects;
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    }

    @SuppressWarnings("resource")
    public <T> List<T> read(File file, String sheetName,
                            ExcelRowMapper<T> rowMapper, boolean ignoreFirstRow)
            throws EncryptedDocumentException, InvalidFormatException,
            IOException {
        Workbook workbook = null;
        try {
            workbook = openWorkbook(file);
            Sheet sheet = workbook.getSheet(sheetName);
            List<T> objects = new ArrayList<>();
            for (Row row : sheet) {
                if (row.getRowNum() > 0 || !ignoreFirstRow) {
                    objects.add(rowMapper.map(row));
                }
            }
            return objects;
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    }

    @SuppressWarnings("resource")
    public <T> List<T> read(InputStream inputStream, String sheetName,
                            ExcelRowMapper<T> rowMapper, boolean ignoreFirstRow)
            throws Exception {
        Workbook workbook = null;
        try {
            workbook = openWorkbook(inputStream);
            Sheet sheet = workbook.getSheet(sheetName);
            List<T> objects = new ArrayList<>();
            for (Row row : sheet) {
                if (row.getRowNum() > 0 || !ignoreFirstRow) {
                    objects.add(rowMapper.map(row));
                }
            }
            return objects;
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    }

    @SuppressWarnings("resource")
    public <T> List<T> read(File file, int sheetIndex,
                            ExcelRowMapperWithHeader<T> rowMapper)
            throws EncryptedDocumentException, InvalidFormatException,
            IOException {
        Workbook workbook = null;
        try {
            workbook = openWorkbook(file);
            Sheet sheet = workbook.getSheetAt(sheetIndex);
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
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    }

    @SuppressWarnings("resource")
    public <T> List<T> read(InputStream inputStream, int sheetIndex,
                            ExcelRowMapperWithHeader<T> rowMapper) throws Exception {
        Workbook workbook = null;
        try {
            workbook = openWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(sheetIndex);
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
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    }

    @SuppressWarnings("resource")
    public <T> List<T> read(File file, String sheetName,
                            ExcelRowMapperWithHeader<T> rowMapper)
            throws EncryptedDocumentException, InvalidFormatException,
            IOException {
        Workbook workbook = null;
        try {
            workbook = openWorkbook(file);
            Sheet sheet = workbook.getSheet(sheetName);
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
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    }

    @SuppressWarnings("resource")
    public <T> List<T> read(InputStream inputStream, String sheetName,
                            ExcelRowMapperWithHeader<T> rowMapper) throws Exception {
        Workbook workbook = null;
        try {
            workbook = openWorkbook(inputStream);
            Sheet sheet = workbook.getSheet(sheetName);
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
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    }

    public Workbook openWorkbook(File file) throws EncryptedDocumentException,
            InvalidFormatException, IOException {
        return WorkbookFactory.create(file);
    }

    public Workbook openWorkbook(InputStream inputStream)
            throws EncryptedDocumentException, InvalidFormatException,
            IOException {
        return WorkbookFactory.create(inputStream);
    }

    private Map<String, Integer> mapHeaders(Row row) {
        ExcelRowMapper<Map<String, Integer>> mapper = DEFAULT_HEADER_MAPPER;
        if (headerMapper != null) {
            mapper = headerMapper;
        }
        return mapper.map(row);
    }
}
