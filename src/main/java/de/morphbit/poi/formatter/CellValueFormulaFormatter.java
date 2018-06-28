package de.morphbit.poi.formatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class CellValueFormulaFormatter extends CellValueFormatter<String> {

	private static CellValueFormulaFormatter instance =
	        new CellValueFormulaFormatter();
	
	private CellValueFormulaFormatter() {
	}
	
	public static CellValueFormulaFormatter getInstance() {
		if(instance == null) {
			instance = new CellValueFormulaFormatter();
		}
		return instance;
	}
	
	@Override
	public String format(Cell cell) {
		if(cell.getCellTypeEnum() == CellType.FORMULA) {
			return cell.getCellFormula();
		}
		return null;
	}

}
