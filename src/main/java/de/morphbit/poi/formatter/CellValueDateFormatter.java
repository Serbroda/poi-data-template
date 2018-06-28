package de.morphbit.poi.formatter;

import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

public class CellValueDateFormatter extends CellValueFormatter<Date> {

	private static CellValueDateFormatter instance =
	        new CellValueDateFormatter();
	
	private CellValueDateFormatter() {
	}
	
	public static CellValueDateFormatter getInstance() {
		if(instance == null) {
			instance = new CellValueDateFormatter();
		}
		return instance;
	}
	
	@Override
	public Date format(Cell cell) {
		if(cell == null) {
			return null;
		}
		if(DateUtil.isCellDateFormatted(cell)) {
			return cell.getDateCellValue();
		}
		return null;
	}

}
