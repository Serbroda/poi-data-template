package de.morphbit.poi;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatCode;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Ignore;
import org.junit.Test;

import de.morphbit.poi.exception.ExcelSourceNotSupportedException;
import de.morphbit.poi.model.ExcelFileSource;
import de.morphbit.poi.model.ExcelSource;

public class ExcelSourceTest extends AbstractResourceTest {

    private static final String FILE_PATH1 = "test1.xlsx";
    private static final String FILE_PATH2 = "test2.xlsx";

    @Test
    public void itShouldOpenFile() throws InvalidFormatException, ExcelSourceNotSupportedException, IOException {
        ExcelSource source = new ExcelFileSource(getResourceFile(FILE_PATH1));
        Workbook workbook = source.openWorkbook();
        assertThat(workbook).isNotNull();
    }

    @Test
    @Ignore
    public void itShouldFailForEncryptedFile() {
        ExcelSource source = new ExcelFileSource(getResourceFile(FILE_PATH2));
        assertThatThrownBy(() -> source.openWorkbook())
                .isInstanceOf(EncryptedDocumentException.class)
                .hasMessageContaining("Export Restrictions");
    }

    @Ignore
    @Test
    public void itShouldNotFailForEncryptedFileWithPasswordSet() {
        ExcelSource source = new ExcelFileSource(getResourceFile(FILE_PATH2), "poidatatemplate");
        assertThatCode(() -> source.openWorkbook()).doesNotThrowAnyException();
    }
}
