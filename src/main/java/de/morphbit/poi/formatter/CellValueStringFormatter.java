package de.morphbit.poi.formatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;

public class CellValueStringFormatter extends CellValueFormatter<String> {

	private final DataFormatter formatter;
	
	public CellValueStringFormatter() {
		this.formatter = new DataFormatter();
	}
	
	@Override
	public String format(Cell cell) {
		if(cell == null) {
			return null;
		}
		if(getEffectiveCellType(cell) == CellType.STRING) {
			return cell.getStringCellValue();
		}
		return formatter.formatCellValue(cell);
	}

}
