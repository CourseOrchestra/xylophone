package ru.curs.xylophone;

import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import static org.junit.Assert.assertTrue;

public class WriterODSTest {

    @Test
    public void WriterODS() throws ODS2SpreadSheetError, IOException {
        InputStream templateStream = TestReader.class
                .getResourceAsStream("template.ods");

        new ODSReportWriter(templateStream, templateStream);

        assertTrue(true);
    }

}
