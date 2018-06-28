package de.morphbit.poi.formatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class CellValueIntegerFormatter extends CellValueFormatter<Integer> {

	private static CellValueIntegerFormatter instance =
	        new CellValueIntegerFormatter();
	
	private CellValueIntegerFormatter() {
	}
	
	public static CellValueIntegerFormatter getInstance() {
		if(instance == null) {
			instance = new CellValueIntegerFormatter();
		}
		return instance;
	}
	
	@Override
	public Integer format(Cell cell) {
		if(cell == null) {
			return null;
		}
		CellType cellType = getEffectiveCellType(cell);
		if(cellType == CellType.NUMERIC) {
			return (int)cell.getNumericCellValue();
		} else if(cellType == CellType.STRING && cell.getStringCellValue().matches("^\\d*$")) {
			return Integer.parseInt(cell.getStringCellValue());
		}
		return null;
	}

}
