package de.morphbit.poi.formatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class CellValueDoubleFormatter extends CellValueFormatter<Double> {

	private static CellValueDoubleFormatter instance =
	        new CellValueDoubleFormatter();
	
	private CellValueDoubleFormatter() {
	}
	
	public static CellValueDoubleFormatter getInstance() {
		if(instance == null) {
			instance = new CellValueDoubleFormatter();
		}
		return instance;
	}
	
	@Override
	public Double format(Cell cell) {
		CellType cellType = getEffectiveCellType(cell);
		if(cellType == CellType.NUMERIC) {
			return cell.getNumericCellValue();
		} else if(cellType == CellType.STRING && cell.getStringCellValue().matches("^\\d*[,|.]{0,1}\\d*$")) {
			return Double.parseDouble(cell.getStringCellValue());
		}
		return null;
	}

}
