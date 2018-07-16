package de.morphbit.poi.mapper;

import org.apache.poi.ss.usermodel.Row;

public interface ExcelDataMapper<T> {

	void map(Row row, T data);
}
