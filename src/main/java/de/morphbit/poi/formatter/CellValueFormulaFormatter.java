package de.morphbit.poi.formatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class CellValueFormulaFormatter extends CellValueFormatter<String> {

	@Override
	public String format(Cell cell) {
		if(cell == null) {
			return null;
		}
		if(cell.getCellTypeEnum() == CellType.FORMULA) {
			return cell.getCellFormula();
		}
		return null;
	}

}
