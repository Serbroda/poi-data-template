package de.morphbit.poi.mapper;

import org.apache.poi.ss.usermodel.Row;

import java.util.Map;

public interface ExcelRowMapperWithHeader<T> {

    T map(Row row, Map<String, Integer> headers);

}
