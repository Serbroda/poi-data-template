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
		Cell cell = createCell();
		cell.setCellType(CellType.STRING);
		cell.setCellValue("Test");
		
		String formattedValue = CellUtils.formatCellAsString(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo("Test");
		
		cell = createCell();
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(874);
		
		formattedValue = CellUtils.formatCellAsString(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo("874");
		
		cell = createCell();
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(1337.23);
		
		formattedValue = CellUtils.formatCellAsString(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo("1337,23");
	}
	
	@Test
	public void itShouldFormatInteger() {
		Cell cell = createCell();
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(1234);
		
		Integer formattedValue = CellUtils.formatCellAsInteger(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo(1234);
		
		cell = createCell();
		cell.setCellType(CellType.STRING);
		cell.setCellValue("874");
		
		formattedValue = CellUtils.formatCellAsInteger(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo(874);
	}
	
	@Test
	public void itShouldFormatDouble() {
		Cell cell = createCell();
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(1234.32);
		
		Double formattedValue = CellUtils.formatCellAsDouble(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo(1234.32);
		
		cell = createCell();
		cell.setCellType(CellType.STRING);
		cell.setCellValue("874.12");
		
		formattedValue = CellUtils.formatCellAsDouble(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo(874.12);
		
		cell = createCell();
		cell.setCellType(CellType.STRING);
		cell.setCellValue("50");
		
		formattedValue = CellUtils.formatCellAsDouble(cell);
		assertThat(formattedValue).isNotNull();
		assertThat(formattedValue).isEqualTo(50);
	}
	
	private Cell createCell() {
		Cell cell = row.createCell(cellCounter);
		cellCounter++;
		return cell;
	}
	
}
