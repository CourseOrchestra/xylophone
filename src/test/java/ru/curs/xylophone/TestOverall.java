package ru.curs.xylophone;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.approvaltests.Approvals;
import org.approvaltests.writers.ApprovalBinaryFileWriter;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.nio.file.Paths;

public class TestOverall {
    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

	@Test
	public void test1() throws XML2SpreadSheetError {
		InputStream descrStream = TestReader.class
				.getResourceAsStream("testdescriptor3.xml");
		InputStream dataStream = TestReader.class
				.getResourceAsStream("testdata.xml");
		InputStream templateStream = TestReader.class
				.getResourceAsStream("template.xls");

		// write results to binary buffer
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		XML2Spreadsheet.process(dataStream, descrStream, templateStream,
				OutputType.XLS, false, bos);
		byte[] writtenData = bos.toByteArray();

		// verify it
		Approvals.verify(new ApprovalBinaryFileWriter(
				new ByteArrayInputStream(writtenData), "xls"
		));
	}

	@Test
	public void test2() throws XML2SpreadSheetError {
		InputStream descrStream = TestReader.class
				.getResourceAsStream("testsaxdescriptor3.xml");
		InputStream dataStream = TestReader.class
				.getResourceAsStream("testdata.xml");
		InputStream templateStream = TestReader.class
				.getResourceAsStream("template.xls");

		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		XML2Spreadsheet.process(dataStream, descrStream, templateStream,
				OutputType.XLS, true, bos);
		byte[] writtenData = bos.toByteArray();

		Approvals.verify(new ApprovalBinaryFileWriter(
				new ByteArrayInputStream(writtenData), "xls"
		));
	}

	@Test
	public void checkGenerateResultXlsFileWithSpecialSymbolsInDataXmlShouldSucceed() throws Exception {
	    File descriptor = Paths.get(TestOverall.class.getResource("testdescriptor.xml").toURI()).toFile();
	    InputStream dataStream = TestReader.class
				.getResourceAsStream("test_data_with_spec_symbols.xml");
	    File template = Paths.get(TestOverall.class.getResource("template.xls").toURI()).toFile();


		ByteArrayOutputStream excel_bos = new ByteArrayOutputStream();
		XML2Spreadsheet.process(dataStream, descriptor, template, false, false, excel_bos);
		byte[] writtenXLSData = excel_bos.toByteArray();

		// verifying the XLS file
		Approvals.verify(new ApprovalBinaryFileWriter(
				new ByteArrayInputStream(writtenXLSData), "xls"
		));

        HSSFWorkbook workbook = new HSSFWorkbook(new ByteArrayInputStream(writtenXLSData));
        Excel2Print excelPrinter = new Excel2Print(workbook);
        excelPrinter.setFopConfig(Paths.get(TestFO.class.getResource("fop.xconf").toURI()).toFile());

		ByteArrayOutputStream pdf_bos = new ByteArrayOutputStream();
        excelPrinter.toPDF(pdf_bos);
		byte[] writtenPDFData = pdf_bos.toByteArray();

		// verifying the PDF file
		Approvals.verify(new ApprovalBinaryFileWriter(
				new ByteArrayInputStream(writtenXLSData), "pdf"
		));
    }
}
