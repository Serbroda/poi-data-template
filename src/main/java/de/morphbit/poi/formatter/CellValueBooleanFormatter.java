package de.morphbit.poi.formatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class CellValueBooleanFormatter extends CellValueFormatter<Boolean> {

	public static final String DEFAULT_FALSE_PATTERN = "^(0)|(false)|(nein)|(falsch)|(no)|(n)$";
	public static final String DEFAULT_TRUE_PATTERN = "^(1)|(true)|(ja)|(richtig)|(yes)|(y)|(j)|(wahr)$";

	private String falsePattern;
	private String truePattern;

	public CellValueBooleanFormatter() {
		this(DEFAULT_FALSE_PATTERN, DEFAULT_TRUE_PATTERN);
	}

	public CellValueBooleanFormatter(String falsePattern, String truePattern) {
		this.falsePattern = falsePattern;
		this.truePattern = truePattern;
	}

	public String getFalsePattern() {
		return falsePattern;
	}

	public void setFalsePattern(String falsePattern) {
		this.falsePattern = falsePattern;
	}

	public String getTruePattern() {
		return truePattern;
	}

	public void setTruePattern(String truePattern) {
		this.truePattern = truePattern;
	}

	@Override
	public Boolean format(Cell cell) {
		if (cell == null) {
			return null;
		}
		CellType type = getEffectiveCellType(cell);
		if (type == CellType.BOOLEAN) {
			return cell.getBooleanCellValue();
		} else if (type == CellType.NUMERIC) {
			if (cell.getNumericCellValue() == 0) {
				return false;
			} else if (cell.getNumericCellValue() == 1) {
				return true;
			} else {
				return null;
			}
		} else if (type == CellType.STRING) {
			String value = cell.getStringCellValue().toLowerCase();
			if (value.matches(getFalsePattern())) {
				return false;
			} else if (value.matches(getTruePattern())) {
				return true;
			} else {
				return null;
			}
		}

		return null;
	}

}
