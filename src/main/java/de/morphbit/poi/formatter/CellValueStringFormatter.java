package de.morphbit.poi.formatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;

public class CellValueStringFormatter extends CellValueFormatter<String> {

	private static CellValueStringFormatter instance =
	        new CellValueStringFormatter();
	
	private final DataFormatter formatter;
	
	private CellValueStringFormatter() {
		this.formatter = new DataFormatter();
	}
	
	public static CellValueStringFormatter getInstance() {
		if(instance == null) {
			instance = new CellValueStringFormatter();
		}
		return instance;
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
