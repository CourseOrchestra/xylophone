package ru.curs.xylophone;

import org.junit.Test;

import java.io.InputStream;
import java.io.OutputStream;

import static org.junit.Assert.assertTrue;

public class WriterODSTest {

    @Test
    public void WriterODS() throws XML2SpreadSheetError, Exception {
        InputStream descrStream = TestReader.class
                .getResourceAsStream("testdescriptor3.xml");
        InputStream dataStream = TestReader.class
                .getResourceAsStream("testdata.xml");
        InputStream templateStream = TestReader.class
                .getResourceAsStream("template.xls");

        new ODSReportWriter(templateStream, templateStream);

//        XML2SpreadseetBLOB b = new XML2SpreadseetBLOB();
//        OutputStream fos = b.getOutStream();
//        XML2Spreadsheet.process(dataStream, descrStream, templateStream,
//                OutputType.ODS, false, fos);
//
//        assertTrue(b.size() > 6000);
    }

}
