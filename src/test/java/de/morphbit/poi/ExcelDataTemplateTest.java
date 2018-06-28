package de.morphbit.poi;

import de.morphbit.poi.exception.ExcelReadException;
import de.morphbit.poi.model.ExcelFileSource;
import de.morphbit.poi.model.ExcelSource;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

public class ExcelDataTemplateTest extends AbstractResourceTest {

    private static final String FILE_PATH1 = "test1.xlsx";

    private ExcelSource testSource1;

    @Before
    public void setUp() {
        testSource1 = new ExcelFileSource(getResourceFile(FILE_PATH1));
    }

    @Test
    public void itShouldReadList() throws IOException, ExcelReadException {
        List<Data> data = new ExcelDataTemplate()
                .read(testSource1, 0, row -> {
                    Data d = new Data();
                    d.setId((int) row.getCell(0).getNumericCellValue());
                    d.setFirstName(row.getCell(1).getStringCellValue());
                    d.setLastName(row.getCell(2).getStringCellValue());
                    d.setDate(row.getCell(3).getDateCellValue());
                    d.setSales(
                            new BigDecimal(row.getCell(4).getNumericCellValue()));
                    return d;
                }, true);

        assertThat(data).isNotNull();
        assertThat(data).hasSize(2);
        assertThat(data.get(0).getId()).isEqualTo(1);
        assertThat(data.get(1).getId()).isEqualTo(2);
        assertThat(data.get(0).getFirstName()).isEqualTo("Max");
        assertThat(data.get(0).getLastName()).isEqualTo("Mustermann");
    }
    
    @Test
    public void itShouldReadFirstLineAsHeader() throws IOException, ExcelReadException {
        Map<String, Integer> headers = new ExcelDataTemplate()
                .readFirstLineAsHeader(testSource1, 0);

        assertThat(headers).isNotNull();
        assertThat(headers).hasSize(5);
        assertThat(headers.containsKey("ID")).isTrue();
        assertThat(headers.containsKey("FIRST_NAME")).isTrue();
        assertThat(headers.containsKey("LAST_NAME")).isTrue();
        assertThat(headers.containsKey("DATE")).isTrue();
        assertThat(headers.containsKey("SALES")).isTrue();
    }


    public static class Data {

        private int id;
        private String firstName;
        private String lastName;
        private Date date;
        private BigDecimal sales;

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
            return "Data [id=" + id + ", firstName=" + firstName + ", lastName="
                    + lastName + ", date=" + date + ", sales=" + sales + "]";
        }

    }
}
