package de.morphbit.poi;

import static org.assertj.core.api.Assertions.assertThat;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;

import de.morphbit.poi.utils.CellUtils;

public class CellValueFormatterTest {

	private Row row;
	private int cellCounter = 0;
	
	@SuppressWarnings("resource")
	@Before
	public void setUp() {
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		row = sheet.createRow(0);
		cellCounter = 0;
	}
	
	@Test
	public void itShouldFormatString() {
		Cell cell = createCell("Test");
		String formattedValue = CellUtils.valueAsString(cell);
		assertThat(formattedValue).isEqualTo("Test");
		
		cell = createCell(874);
		formattedValue = CellUtils.valueAsString(cell);
		assertThat(formattedValue).isEqualTo("874");
		
		cell = createCell(1337.23);
		formattedValue = CellUtils.valueAsString(cell);
		assertThat(formattedValue).isEqualTo("1337,23");
	}
	
	@Test
	public void itShouldFormatInteger() {
		Cell cell = createCell(1234);
		Integer formattedValue = CellUtils.valueAsInteger(cell);
		assertThat(formattedValue).isEqualTo(1234);
		
		cell = createCell("874");
		formattedValue = CellUtils.valueAsInteger(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo(874);
	}
	
	@Test
	public void itShouldFormatDouble() {
		Cell cell = createCell(1234.32);
		Double formattedValue = CellUtils.valueAsDouble(cell);
		assertThat(formattedValue).isEqualTo(1234.32);
		
		cell = createCell("874.12");
		formattedValue = CellUtils.valueAsDouble(cell);
		assertThat(formattedValue).isEqualTo(874.12);
		
		cell = createCell("50");
		formattedValue = CellUtils.valueAsDouble(cell);
		assertThat(formattedValue).isEqualTo(50);
	}
	
	private Cell createCell() {
		Cell cell = row.createCell(cellCounter);
		cellCounter++;
		return cell;
	}
	
	private Cell createCell(Object value) {
		Cell cell = createCell();
		if (value instanceof String) {
			cell.setCellType(CellType.STRING);
			cell.setCellValue((String) value);
		} else if (value instanceof Integer) {
			cell.setCellType(CellType.NUMERIC);
			cell.setCellValue((Integer) value);
		} else if (value instanceof Double) {
			cell.setCellType(CellType.NUMERIC);
			cell.setCellValue((Double) value);
		} else if (value instanceof Long) {
			cell.setCellType(CellType.NUMERIC);
			cell.setCellValue((Long) value);
		} else if (value instanceof Boolean) {
			cell.setCellType(CellType.BOOLEAN);
			cell.setCellValue((Boolean) value);
		}
		return cell;
	}
	
}
