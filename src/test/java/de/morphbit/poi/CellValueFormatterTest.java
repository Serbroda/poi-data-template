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
		Cell cell = createCell();
		cell.setCellType(CellType.STRING);
		cell.setCellValue("Test");
		
		String formattedValue = CellValueUtils.formatCellAsString(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo("Test");
		
		cell = createCell();
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(874);
		
		formattedValue = CellValueUtils.formatCellAsString(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo("874");
		
		cell = createCell();
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(1337.23);
		
		formattedValue = CellValueUtils.formatCellAsString(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo("1337,23");
	}
	
	private Cell createCell() {
		Cell cell = row.createCell(cellCounter);
		cellCounter++;
		return cell;
	}
	
}
