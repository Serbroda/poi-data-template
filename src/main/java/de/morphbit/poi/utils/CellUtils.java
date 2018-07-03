package de.morphbit.poi.utils;

import java.math.BigDecimal;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;

import de.morphbit.poi.formatter.CellValueBigDecimalFormatter;
import de.morphbit.poi.formatter.CellValueBooleanFormatter;
import de.morphbit.poi.formatter.CellValueDateFormatter;
import de.morphbit.poi.formatter.CellValueDoubleFormatter;
import de.morphbit.poi.formatter.CellValueFormatter;
import de.morphbit.poi.formatter.CellValueFormulaFormatter;
import de.morphbit.poi.formatter.CellValueIntegerFormatter;
import de.morphbit.poi.formatter.CellValueStringFormatter;

public final class CellUtils {

	private static final CellValueFormatter<Boolean> BOOLEAN_FORMATTER = new CellValueBooleanFormatter();
	private static final CellValueFormatter<String> STRING_FORMATTER = new CellValueStringFormatter();
	private static final CellValueFormatter<Integer> INTEGER_FORMATTER = new CellValueIntegerFormatter();
	private static final CellValueFormatter<Date> DATE_FORMATTER = new CellValueDateFormatter();
	private static final CellValueFormatter<BigDecimal> BIGDECIMAL_FORMATTER = new CellValueBigDecimalFormatter();
	private static final CellValueFormatter<Double> DOUBLE_FORMATTER = new CellValueDoubleFormatter();
	private static final CellValueFormatter<String> FORMULA_FORMATTER = new CellValueFormulaFormatter();
	
	private CellUtils() {
		
	}
	
	public static String valueAsString(Cell cell) {
		return value(STRING_FORMATTER, cell);
	}
	
	public static String valueAsString(Cell cell, String defaultValue) {
		return value(STRING_FORMATTER, cell, defaultValue);
	}
	
	public static Boolean valueAsBoolean(Cell cell) {
		return value(BOOLEAN_FORMATTER, cell);
	}
	
	public static Boolean valueAsBoolean(Cell cell, Boolean defaultValue) {
		return value(BOOLEAN_FORMATTER, cell, defaultValue);
	}
	
	public static Date valueAsDate(Cell cell) {
		return value(DATE_FORMATTER, cell);
	}
	
	public static Date valueAsDate(Cell cell, Date defaultValue) {
		return value(DATE_FORMATTER, cell, defaultValue);
	}
	
	public static Integer valueAsInteger(Cell cell) {
		return value(INTEGER_FORMATTER, cell);
	}
	
	public static Integer valueAsInteger(Cell cell, Integer defaultValue) {
		return value(INTEGER_FORMATTER, cell, defaultValue);
	}
	
	public static Double valueAsDouble(Cell cell) {
		return value(DOUBLE_FORMATTER, cell);
	}
	
	public static Double valueAsDouble(Cell cell, Double defaultValue) {
		return value(DOUBLE_FORMATTER, cell, defaultValue);
	}
	
	public static BigDecimal valueAsBigDecimal(Cell cell) {
		return value(BIGDECIMAL_FORMATTER, cell);
	}
	
	public static BigDecimal valueAsBigDecimal(Cell cell, BigDecimal defaultValue) {
		return value(BIGDECIMAL_FORMATTER, cell, defaultValue);
	}
	
	public static String valueAsFormula(Cell cell) {
		return value(FORMULA_FORMATTER, cell);
	}
	
	public static String valueAsFormula(Cell cell, String defaultValue) {
		return value(FORMULA_FORMATTER, cell, defaultValue);
	}
	
	public static <T> T value(CellValueFormatter<T> formatter, Cell cell) {
		return formatter.format(cell);
	}
	
	public static <T> T value(CellValueFormatter<T> formatter, Cell cell, T defaultValue) {
		return formatter.format(cell, defaultValue);
	}
}
