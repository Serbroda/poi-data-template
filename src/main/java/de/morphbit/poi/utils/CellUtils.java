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
	
	public static String formatCellAsString(Cell cell) {
		return formatCell(STRING_FORMATTER, cell);
	}
	
	public static String formatCellAsString(Cell cell, String defaultValue) {
		return formatCell(STRING_FORMATTER, cell, defaultValue);
	}
	
	public static Boolean formatCellAsBoolean(Cell cell) {
		return formatCell(BOOLEAN_FORMATTER, cell);
	}
	
	public static Boolean formatCellAsBoolean(Cell cell, Boolean defaultValue) {
		return formatCell(BOOLEAN_FORMATTER, cell, defaultValue);
	}
	
	public static Date formatCellAsDate(Cell cell) {
		return formatCell(DATE_FORMATTER, cell);
	}
	
	public static Date formatCellAsDate(Cell cell, Date defaultValue) {
		return formatCell(DATE_FORMATTER, cell, defaultValue);
	}
	
	public static Integer formatCellAsInteger(Cell cell) {
		return formatCell(INTEGER_FORMATTER, cell);
	}
	
	public static Integer formatCellAsInteger(Cell cell, Integer defaultValue) {
		return formatCell(INTEGER_FORMATTER, cell, defaultValue);
	}
	
	public static Double formatCellAsDouble(Cell cell) {
		return formatCell(DOUBLE_FORMATTER, cell);
	}
	
	public static Double formatCellAsDouble(Cell cell, Double defaultValue) {
		return formatCell(DOUBLE_FORMATTER, cell, defaultValue);
	}
	
	public static BigDecimal formatCellAsBigDecimal(Cell cell) {
		return formatCell(BIGDECIMAL_FORMATTER, cell);
	}
	
	public static BigDecimal formatCellAsBigDecimal(Cell cell, BigDecimal defaultValue) {
		return formatCell(BIGDECIMAL_FORMATTER, cell, defaultValue);
	}
	
	public static String formatCellAsFormula(Cell cell) {
		return formatCell(FORMULA_FORMATTER, cell);
	}
	
	public static String formatCellAsFormula(Cell cell, String defaultValue) {
		return formatCell(FORMULA_FORMATTER, cell, defaultValue);
	}
	
	public static <T> T formatCell(CellValueFormatter<T> formatter, Cell cell) {
		return formatter.format(cell);
	}
	
	public static <T> T formatCell(CellValueFormatter<T> formatter, Cell cell, T defaultValue) {
		return formatter.format(cell, defaultValue);
	}
}
