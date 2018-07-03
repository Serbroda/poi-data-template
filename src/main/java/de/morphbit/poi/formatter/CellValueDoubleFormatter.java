package de.morphbit.poi.formatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class CellValueDoubleFormatter extends CellValueFormatter<Double> {

	@Override
	public Double format(Cell cell) {
		if(cell == null) {
			return null;
		}
		CellType cellType = getEffectiveCellType(cell);
		if(cellType == CellType.NUMERIC) {
			return cell.getNumericCellValue();
		} else if(cellType == CellType.STRING && cell.getStringCellValue().matches("^\\d*[,|.]{0,1}\\d*$")) {
			return Double.parseDouble(cell.getStringCellValue());
		}
		return null;
	}

}
