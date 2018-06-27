package de.morphbit.poi.mapper;

import org.apache.poi.ss.usermodel.Row;

public interface ExcelRowMapper<T> {

    T map(Row row);
}
