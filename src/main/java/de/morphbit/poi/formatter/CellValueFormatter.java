package de.morphbit.poi.formatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public abstract class CellValueFormatter<T> {

	public abstract T format(Cell cell);
	
	public T format(Cell cell, T defaultValue) {
		T val = format(cell);
		if(val == null) {
			val = defaultValue;
		}
		return val;
	}
	
	public CellType getEffectiveCellType(Cell cell) {
		if(cell == null) {
			return null;
		}
		if(isCellFormula(cell)) {
			return cell.getCachedFormulaResultTypeEnum();
		}
		return cell.getCellTypeEnum();
	}
	
	public boolean isCellFormula(Cell cell) {
		if(cell == null || cell.getCellTypeEnum() == null) {
			return false;
		}
		return cell.getCellTypeEnum() == CellType.FORMULA;
	}
	
}
