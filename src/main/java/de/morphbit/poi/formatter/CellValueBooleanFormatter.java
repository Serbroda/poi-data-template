package de.morphbit.poi.formatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class CellValueBooleanFormatter extends CellValueFormatter<Boolean> {

	private static CellValueBooleanFormatter instance;
	
	public static final String DEFAULT_FALSE_REGEX = "^(0)|(false)|(nein)|(falsch)|(no)|(n)$";
	public static final String DEFAULT_TRUE_REGEX = "^(1)|(true)|(ja)|(richtig)|(yes)|(y)|(j)|(wahr)$";
	
	public static CellValueBooleanFormatter getInstance() {
		if(instance == null) {
			instance = new CellValueBooleanFormatter();
		}
		return instance;
	}
	
	@Override
	public Boolean format(Cell cell) {
		if(cell == null) {
			return null;
		}
		CellType type = getEffectiveCellType(cell);
		if(type == CellType.BOOLEAN) {
			return cell.getBooleanCellValue();
		 } else if (type == CellType.NUMERIC) {
			 if(cell.getNumericCellValue() == 0) {
				 return false;
			 } else if (cell.getNumericCellValue() == 1) {
				 return true;
			 } else {
				 return null;
			 }
		 } else if (type == CellType.STRING) {
			 String value = cell.getStringCellValue().toLowerCase();
			 if(value.matches(DEFAULT_FALSE_REGEX)) {
				 return false;
			 } else if (value.matches(DEFAULT_TRUE_REGEX)) {
				 return true;
			 } else {
				 return null;
			 }
		 }
		
		return null;
	}

}
