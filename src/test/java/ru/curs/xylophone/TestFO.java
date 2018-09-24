package ru.curs.xylophone;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.*;
import java.nio.file.Paths;

import static org.junit.Assert.assertTrue;

public class TestFO {
    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void test1() throws Exception {
        try (InputStream is = TestFO.class.getResourceAsStream("template.xls");
             OutputStream os = new ByteArrayOutputStream()
             // os = new FileOutputStream("d:/temp/xls.pdf");
        ) {
            Excel2Print e2p = new Excel2Print(is);
            /*
             * e2p.setFopConfig( new File(
             * "D:/workspaces/workspace/xylophone/src/test/resources/ru/curs/xylophone/fop.xconf"
             * ));
             */
            e2p.toFO(os);

            // e2p.toPDF(os);
        }
    }

    @Test
    public void convertToPdfWithSymbolOfRoubleAndPercentShouldSuccess() throws Exception {
        File createdFile = temporaryFolder.newFile("after_conversion_to_pdf.pdf");
        try (InputStream is = TestFO.class.getResourceAsStream("before_conversion_to_pdf.xls");
             OutputStream os = new FileOutputStream(createdFile)
        ) {
            Excel2Print e2p = new Excel2Print(is);
             e2p.setFopConfig(Paths.get(TestFO.class.getResource("fop.xconf").toURI()).toFile());

             e2p.toPDF(os);
        }
        assertTrue(createdFile.exists());
        assertTrue(createdFile.length() > 0);
    }
}
