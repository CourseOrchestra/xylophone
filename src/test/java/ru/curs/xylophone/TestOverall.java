package ru.curs.xylophone;

import static org.junit.Assert.assertTrue;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.junit.Test;

public class TestOverall {
	@Test
	public void test1() throws XML2SpreadSheetError, IOException {
		InputStream descrStream = TestReader.class
				.getResourceAsStream("testdescriptor3.xml");
		InputStream dataStream = TestReader.class
				.getResourceAsStream("testdata.xml");
		InputStream templateStream = TestReader.class
				.getResourceAsStream("template.xls");

		XML2SpreadseetBLOB b = new XML2SpreadseetBLOB();
		OutputStream fos = b.getOutStream();
		XML2Spreadsheet.process(dataStream, descrStream, templateStream,
				OutputType.XLS, false, fos);

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
	public void test2() throws XML2SpreadSheetError, IOException {
		InputStream descrStream = TestReader.class
				.getResourceAsStream("testsaxdescriptor3.xml");
		InputStream dataStream = TestReader.class
				.getResourceAsStream("testdata.xml");
		InputStream templateStream = TestReader.class
				.getResourceAsStream("template.xls");

		XML2SpreadseetBLOB b = new XML2SpreadseetBLOB();
		OutputStream fos = b.getOutStream();
		XML2Spreadsheet.process(dataStream, descrStream, templateStream,
				OutputType.XLS, true, fos);
		assertTrue(b.size() > 6000);
		/*
		 * byte[] buffer = new byte[1024]; FileOutputStream out = new
		 * FileOutputStream(new File( "c:/temp/!!!test2.xls")); try {
		 * InputStream in = b.getInStream(); int len = in.read(buffer); while
		 * (len != -1) { out.write(buffer, 0, len); len = in.read(buffer); } }
		 * finally { out.close(); }
		 */
	}
}
