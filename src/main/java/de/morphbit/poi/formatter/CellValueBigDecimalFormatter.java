package de.morphbit.poi.formatter;

import java.math.BigDecimal;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class CellValueBigDecimalFormatter extends CellValueFormatter<BigDecimal> {

	private static CellValueBigDecimalFormatter instance =
	        new CellValueBigDecimalFormatter();
	
	private CellValueBigDecimalFormatter() {
	}
	
	public static CellValueBigDecimalFormatter getInstance() {
		if(instance == null) {
			instance = new CellValueBigDecimalFormatter();
		}
		return instance;
	}
	
	@Override
	public BigDecimal format(Cell cell) {
		CellType cellType = getEffectiveCellType(cell);
		if(cellType == CellType.NUMERIC) {
			return new BigDecimal(cell.getNumericCellValue());
		} else if(cellType == CellType.STRING && cell.getStringCellValue().matches("^\\d*[,|.]{1}\\d*$")) {
			return new BigDecimal(cell.getStringCellValue());
		}
		return null;
	}

}
