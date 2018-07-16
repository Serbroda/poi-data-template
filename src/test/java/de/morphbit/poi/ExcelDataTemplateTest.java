package de.morphbit.poi;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.function.Predicate;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.junit.Before;
import org.junit.Test;

import de.morphbit.poi.exception.ExcelReadException;
import de.morphbit.poi.mapper.ExcelDataMapper;
import de.morphbit.poi.mapper.ExcelRowMapper;
import de.morphbit.poi.mapper.ExcelRowMapperWithHeader;
import de.morphbit.poi.model.ExcelDataTemplateOptions;
import de.morphbit.poi.model.ExcelFileSource;
import de.morphbit.poi.model.ExcelSource;
import de.morphbit.poi.utils.CellUtils;

public class ExcelDataTemplateTest extends AbstractResourceTest {

	private static final String FILE_PATH1 = "test1.xlsx";

	private ExcelSource testSource1;
	private ExcelRowMapper<Data> simpleRowMapper;

	@Before
	public void setUp() {
		testSource1 = new ExcelFileSource(getResourceFile(FILE_PATH1));
		simpleRowMapper = row -> {
			Data d = new Data();
			d.setId((int) row.getCell(0).getNumericCellValue());
			d.setFirstName(row.getCell(1).getStringCellValue());
			d.setLastName(row.getCell(2).getStringCellValue());
			d.setDate(row.getCell(3).getDateCellValue());
			d.setSales(new BigDecimal(row.getCell(4).getNumericCellValue()));
			return d;
		};
	}

	@Test
	public void itShouldReadList() throws IOException, ExcelReadException {
		List<Data> data = new ExcelDataTemplate().read(testSource1, 0, simpleRowMapper,
				new ExcelDataTemplateOptions().withIgnoreFirstLinesCount(1));

		assertThat(data).isNotNull();
		assertThat(data).hasSize(2);
		assertThat(data.get(0).getId()).isEqualTo(1);
		assertThat(data.get(1).getId()).isEqualTo(2);
		assertThat(data.get(0).getFirstName()).isEqualTo("Max");
		assertThat(data.get(0).getLastName()).isEqualTo("Mustermann");
	}

	@Test
	public void itShouldReadListFormattedWithCellUtils() throws IOException, ExcelReadException {
		List<Data> data = new ExcelDataTemplate().read(testSource1, 0, row -> {
			Data d = new Data();
			d.setId(CellUtils.valueAsInteger(row.getCell(0)));
			d.setFirstName(CellUtils.valueAsString(row.getCell(1)));
			d.setLastName(CellUtils.valueAsString(row.getCell(2)));
			d.setDate(CellUtils.valueAsDate(row.getCell(3)));
			d.setSales(CellUtils.valueAsBigDecimal(row.getCell(4)));
			return d;
		}, new ExcelDataTemplateOptions().withIgnoreFirstLinesCount(1));

		assertThat(data).isNotNull();
		assertThat(data).hasSize(2);
		assertThat(data.get(0).getId()).isEqualTo(1);
		assertThat(data.get(1).getId()).isEqualTo(2);
		assertThat(data.get(0).getFirstName()).isEqualTo("Max");
		assertThat(data.get(0).getLastName()).isEqualTo("Mustermann");
	}

	@Test
	public void itShouldReadListTwiceWithSameSize() throws IOException, ExcelReadException {
		List<Data> data = new ExcelDataTemplate().read(testSource1, 0, simpleRowMapper,
				new ExcelDataTemplateOptions().withIgnoreFirstLinesCount(1));
		assertThat(data).isNotNull();
		assertThat(data).hasSize(2);

		data = new ExcelDataTemplate().read(testSource1, 0, simpleRowMapper,
				new ExcelDataTemplateOptions().withIgnoreFirstLinesCount(1));
		assertThat(data).isNotNull();
		assertThat(data).hasSize(2);
	}

	@Test
	public void itShouldReadRowsWithHeader() throws IOException, ExcelReadException {
		List<Data> data = new ExcelDataTemplate().read(testSource1, 0, new ExcelRowMapperWithHeader<Data>() {

			@Override
			public Data map(Row row, Map<String, Integer> headers) {
				Data d = new Data();
				d.setId((int) row.getCell(headers.get("ID")).getNumericCellValue());
				d.setFirstName(row.getCell(headers.get("FIRST_NAME")).getStringCellValue());
				d.setLastName(row.getCell(headers.get("LAST_NAME")).getStringCellValue());
				d.setDate(row.getCell(headers.get("SALES")).getDateCellValue());
				d.setSales(new BigDecimal(row.getCell(headers.get("ID")).getNumericCellValue()));
				return d;
			}
		});

		assertThat(data).isNotNull();
		assertThat(data).hasSize(2);
		assertThat(data.get(0).getId()).isEqualTo(1);
		assertThat(data.get(1).getId()).isEqualTo(2);
		assertThat(data.get(0).getFirstName()).isEqualTo("Max");
		assertThat(data.get(0).getLastName()).isEqualTo("Mustermann");
	}

	@Test
	public void itShouldReadRowsWithHeaderWithoutOptions() throws IOException, ExcelReadException {
		List<Data> data = new ExcelDataTemplate().read(testSource1, 0, new ExcelRowMapperWithHeader<Data>() {

			@Override
			public Data map(Row row, Map<String, Integer> headers) {
				Data d = new Data();
				d.setId((int) row.getCell(headers.get("ID")).getNumericCellValue());
				d.setFirstName(row.getCell(headers.get("FIRST_NAME")).getStringCellValue());
				d.setLastName(row.getCell(headers.get("LAST_NAME")).getStringCellValue());
				d.setDate(row.getCell(headers.get("SALES")).getDateCellValue());
				d.setSales(new BigDecimal(row.getCell(headers.get("ID")).getNumericCellValue()));
				return d;
			}
		});

		assertThat(data).isNotNull();
	}

	@Test
	public void itShouldFilterList() throws IOException, ExcelReadException {
		ExcelRowMapperWithHeader<Data> rowMapper = new ExcelRowMapperWithHeader<Data>() {

			@Override
			public Data map(Row row, Map<String, Integer> headers) {
				Data d = new Data();
				d.setId((int) row.getCell(headers.get("ID")).getNumericCellValue());
				d.setFirstName(row.getCell(headers.get("FIRST_NAME")).getStringCellValue());
				d.setLastName(row.getCell(headers.get("LAST_NAME")).getStringCellValue());
				d.setDate(row.getCell(headers.get("SALES")).getDateCellValue());
				d.setSales(new BigDecimal(row.getCell(headers.get("ID")).getNumericCellValue()));
				return d;
			}
		};
		List<Data> data = new ExcelDataTemplate().read(testSource1, 0, rowMapper, new ExcelDataTemplateOptions().withFilter(new Predicate<Data>() {

			@Override
			public boolean test(Data t) {
				return t.getId() > 1;
			}
		}));

		assertThat(data).isNotNull();
		assertThat(data).hasSize(1);
		assertThat(data.get(0).getId()).isEqualTo(2);
		assertThat(data.get(0).getFirstName()).isEqualTo("John");
		assertThat(data.get(0).getLastName()).isEqualTo("Doe");
		
		data = new ExcelDataTemplate().read(testSource1, 0, rowMapper, new ExcelDataTemplateOptions().withFilter(new Predicate<Data>() {

			@Override
			public boolean test(Data t) {
				return "Max".equals(t.getFirstName());
			}
		}));

		assertThat(data).isNotNull();
		assertThat(data).hasSize(1);
		assertThat(data.get(0).getId()).isEqualTo(1);
		assertThat(data.get(0).getFirstName()).isEqualTo("Max");
		assertThat(data.get(0).getLastName()).isEqualTo("Mustermann");
		
		data = new ExcelDataTemplate().read(testSource1, 0, rowMapper, new ExcelDataTemplateOptions().withFilter(new Predicate<Data>() {

			@Override
			public boolean test(Data t) {
				return "Unknown".equals(t.getFirstName());
			}
		}));

		assertThat(data).isNotNull();
		assertThat(data).hasSize(0);
	}

	@Test
	public void itShouldReadRowsWithHeaderSeparate() throws IOException, ExcelReadException {
		Map<String, Integer> headers = new ExcelDataTemplate().readFirstLineAsHeader(testSource1, 0);

		List<Data> data = new ExcelDataTemplate().read(testSource1, 0, row -> {
			Data d = new Data();
			d.setId((int) row.getCell(headers.get("ID")).getNumericCellValue());
			d.setFirstName(row.getCell(headers.get("FIRST_NAME")).getStringCellValue());
			d.setLastName(row.getCell(headers.get("LAST_NAME")).getStringCellValue());
			d.setDate(row.getCell(headers.get("SALES")).getDateCellValue());
			d.setSales(new BigDecimal(row.getCell(headers.get("ID")).getNumericCellValue()));
			return d;
		}, new ExcelDataTemplateOptions().withIgnoreFirstLinesCount(1));

		assertThat(data).isNotNull();
		assertThat(data).hasSize(2);
		assertThat(data.get(0).getId()).isEqualTo(1);
		assertThat(data.get(1).getId()).isEqualTo(2);
		assertThat(data.get(0).getFirstName()).isEqualTo("Max");
		assertThat(data.get(0).getLastName()).isEqualTo("Mustermann");
	}

	@Test
	public void itShouldReadFirstLineAsHeader() throws IOException, ExcelReadException {
		Map<String, Integer> headers = new ExcelDataTemplate().readFirstLineAsHeader(testSource1, 0);

		assertThat(headers).isNotNull();
		assertThat(headers).hasSize(5);
		assertThat(headers.containsKey("ID")).isTrue();
		assertThat(headers.containsKey("FIRST_NAME")).isTrue();
		assertThat(headers.containsKey("LAST_NAME")).isTrue();
		assertThat(headers.containsKey("DATE")).isTrue();
		assertThat(headers.containsKey("SALES")).isTrue();
	}
	
	@SuppressWarnings("deprecation")
	@Test
	public void itShouldWriteData() throws IOException, ExcelReadException {
		List<String> headers = new ArrayList<>();
		headers.add("ID");
		headers.add("FIRST_NAME");
		headers.add("LAST_NAME");
		headers.add("DATE");
		headers.add("SALES");
		
		List<Data> data = new ArrayList<>();
		data.add(new Data(1337, "John", "Johnson", new Date(2018, 1, 1), new BigDecimal("12.95")));
		data.add(new Data(4711, "David", "Davidson", new Date(2018, 1, 1), new BigDecimal("12.95")));
		
		new ExcelDataTemplate().write(new File("src/test/resources/test3.xlsx"), headers, data, new ExcelDataMapper<Data>() {

			@Override
			public void map(Row row, Data data) {
				int cellCount = 0;
				
				Cell cell = row.createCell(cellCount++);
				cell.setCellValue(data.getId());
				cell = row.createCell(cellCount++);
				cell.setCellValue(data.getFirstName());
				cell = row.createCell(cellCount++);
				cell.setCellValue(data.getLastName());
				cell = row.createCell(cellCount++);
				cell.setCellValue(data.getDate());
				cell = row.createCell(cellCount++);
				cell.setCellValue(data.getSales().doubleValue());
			}
		});
		
		List<Data> readData = new ExcelDataTemplate().read(new ExcelFileSource(getResourceFile("test3.xlsx")), 0, simpleRowMapper,
				new ExcelDataTemplateOptions().withIgnoreFirstLinesCount(1));

		assertThat(readData).isNotNull();
		assertThat(readData).hasSize(2);
		assertThat(readData.get(0).getId()).isEqualTo(1337);
		assertThat(readData.get(0).getFirstName()).isEqualTo("John");
		assertThat(readData.get(0).getLastName()).isEqualTo("Johnson");
	}

	public static class Data {

		private int id;
		private String firstName;
		private String lastName;
		private Date date;
		private BigDecimal sales;
		
		public Data() {
			
		}

		public Data(int id, String firstName, String lastName, Date date, BigDecimal sales) {
			this.id = id;
			this.firstName = firstName;
			this.lastName = lastName;
			this.date = date;
			this.sales = sales;
		}

		public int getId() {
			return id;
		}

		public void setId(int id) {
			this.id = id;
		}

		public String getFirstName() {
			return firstName;
		}

		public void setFirstName(String firstName) {
			this.firstName = firstName;
		}

		public String getLastName() {
			return lastName;
		}

		public void setLastName(String lastName) {
			this.lastName = lastName;
		}

		public Date getDate() {
			return date;
		}

		public void setDate(Date date) {
			this.date = date;
		}

		public BigDecimal getSales() {
			return sales;
		}

		public void setSales(BigDecimal sales) {
			this.sales = sales;
		}

		@Override
		public String toString() {
			return "Data [id=" + id + ", firstName=" + firstName + ", lastName=" + lastName + ", date=" + date
					+ ", sales=" + sales + "]";
		}

	}
}
