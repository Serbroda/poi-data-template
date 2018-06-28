package de.morphbit.poi.formatter;

import java.math.BigDecimal;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;

public final class CellValueUtils {

	private CellValueUtils() {
		
	}
	
	public static String formatCellAsString(Cell cell) {
		return formatCell(CellValueStringFormatter.getInstance(), cell);
	}
	
	public static String formatCellAsString(Cell cell, String defaultValue) {
		return formatCell(CellValueStringFormatter.getInstance(), cell, defaultValue);
	}
	
	public static Date formatCellAsDate(Cell cell) {
		return formatCell(CellValueDateFormatter.getInstance(), cell);
	}
	
	public static Date formatCellAsDate(Cell cell, Date defaultValue) {
		return formatCell(CellValueDateFormatter.getInstance(), cell, defaultValue);
	}
	
	public static Integer formatCellAsInteger(Cell cell) {
		return formatCell(CellValueIntegerFormatter.getInstance(), cell);
	}
	
	public static Integer formatCellAsInteger(Cell cell, Integer defaultValue) {
		return formatCell(CellValueIntegerFormatter.getInstance(), cell, defaultValue);
	}
	
	public static Double formatCellAsDouble(Cell cell) {
		return formatCell(CellValueDoubleFormatter.getInstance(), cell);
	}
	
	public static Double formatCellAsDouble(Cell cell, Double defaultValue) {
		return formatCell(CellValueDoubleFormatter.getInstance(), cell, defaultValue);
	}
	
	public static BigDecimal formatCellAsBigDecimal(Cell cell) {
		return formatCell(CellValueBigDecimalFormatter.getInstance(), cell);
	}
	
	public static BigDecimal formatCellAsBigDecimal(Cell cell, BigDecimal defaultValue) {
		return formatCell(CellValueBigDecimalFormatter.getInstance(), cell, defaultValue);
	}
	
	public static String formatCellAsFormula(Cell cell) {
		return formatCell(CellValueFormulaFormatter.getInstance(), cell);
	}
	
	public static String formatCellAsFormula(Cell cell, String defaultValue) {
		return formatCell(CellValueFormulaFormatter.getInstance(), cell, defaultValue);
	}
	
	public static <T> T formatCell(CellValueFormatter<T> formatter, Cell cell) {
		return formatter.format(cell);
	}
	
	public static <T> T formatCell(CellValueFormatter<T> formatter, Cell cell, T defaultValue) {
		return formatter.format(cell, defaultValue);
	}
}
