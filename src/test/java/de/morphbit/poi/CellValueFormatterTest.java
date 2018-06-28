package de.morphbit.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;

import de.morphbit.poi.formatter.CellValueUtils;

import static org.assertj.core.api.Assertions.assertThat;

public class CellValueFormatterTest {

	private Row row;
	
	@SuppressWarnings("resource")
	@Before
	public void setUp() {
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		row = sheet.createRow(0);
	}
	
	@Test
	public void itShouldFormatString() {
		Cell cell = row.createCell(0);
		cell.setCellType(CellType.STRING);
		cell.setCellValue("Test");
		
		String formattedValue = CellValueUtils.formatCellAsString(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo("Test");
		
		cell = row.createCell(1);
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(874);
		
		formattedValue = CellValueUtils.formatCellAsString(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo("874");
		
		cell = row.createCell(2);
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(1337.23);
		
		formattedValue = CellValueUtils.formatCellAsString(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo("1337,23");
	}
}
