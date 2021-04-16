package ru.curs.xylophone;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.*;
import java.nio.file.Paths;

import static org.junit.Assert.assertTrue;

public class TestOverall {
    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

	@Test
	public void test1() throws XylophoneError {
		InputStream descrStream = TestReader.class
				.getResourceAsStream("testdescriptor3.json");
		InputStream dataStream = TestReader.class
				.getResourceAsStream("testdata.xml");
		InputStream templateStream = TestReader.class
				.getResourceAsStream("template.xls");

		XML2SpreadseetBLOB b = new XML2SpreadseetBLOB();
		OutputStream fos = b.getOutStream();
		XML2Spreadsheet.process(dataStream, descrStream, templateStream,
				OutputType.XLS, false, false, fos);

		assertTrue(b.size() > 6000);
		/*
		 * byte[] buffer = new byte[1024]; FileOutputStream out = new
		 * FileOutputStream(new File( "c:/temp/!!!test1.xls")); try {
		 * InputStream in = b.getInStream(); int len = in.read(buffer); while
		 * (len != -1) { out.write(buffer, 0, len); len = in.read(buffer); } }
		 * finally { out.close(); }
		 */
	}

	@Test
	public void test2() throws XylophoneError {
		InputStream descrStream = TestReader.class
				.getResourceAsStream("testsaxdescriptor3.json");
		InputStream dataStream = TestReader.class
				.getResourceAsStream("testdata.xml");
		InputStream templateStream = TestReader.class
				.getResourceAsStream("template.xls");

		XML2SpreadseetBLOB b = new XML2SpreadseetBLOB();
		OutputStream fos = b.getOutStream();
		XML2Spreadsheet.process(dataStream, descrStream, templateStream,
				OutputType.XLS, true, false, fos);
		assertTrue(b.size() > 6000);
		/*
		 * byte[] buffer = new byte[1024]; FileOutputStream out = new
		 * FileOutputStream(new File( "c:/temp/!!!test2.xls")); try {
		 * InputStream in = b.getInStream(); int len = in.read(buffer); while
		 * (len != -1) { out.write(buffer, 0, len); len = in.read(buffer); } }
		 * finally { out.close(); }
		 */
	}

	@Test
	public void checkGenerateResultXlsFileWithSpecialSymbolsInDataXmlShouldSuccess() throws Exception {
		File descriptorFile = Paths.get(TestOverall.class.getResource("testdescriptor.json").toURI()).toFile();
		FileInputStream descriptor = new FileInputStream(descriptorFile);
	    InputStream dataStream = TestReader.class
				.getResourceAsStream("test_data_with_spec_symbols.xml");
	    File template = Paths.get(TestOverall.class.getResource("template.xls").toURI()).toFile();

        File createdTempOutputFile = temporaryFolder.newFile("temp.xls");

        try (OutputStream outputStream = new FileOutputStream(createdTempOutputFile)) {
            XML2Spreadsheet.process(dataStream, descriptor, template,
                    false, false, outputStream);
        }

        NPOIFSFileSystem fileSystem = new NPOIFSFileSystem(createdTempOutputFile);
        HSSFWorkbook workbook = new HSSFWorkbook(fileSystem.getRoot(), false);

        Excel2Print excelPrinter = new Excel2Print(workbook);
        excelPrinter.setFopConfig(Paths.get(TestFO.class.getResource("fop.xconf").toURI()).toFile());

        File pdfResultFile = temporaryFolder.newFile("after_conversion_to_pdf.pdf");
        excelPrinter.toPDF(new FileOutputStream(pdfResultFile));

        assertTrue(createdTempOutputFile.length() > 0);
        assertTrue(pdfResultFile.length() > 0);
    }
}
