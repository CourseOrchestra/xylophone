package ru.curs.xylophone;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Sheet;
import org.junit.After;
import org.junit.Rule;
import org.junit.Test;

import org.junit.rules.ExpectedException;
import ru.curs.xylophone.XMLDataReader.DescriptorElement;
import ru.curs.xylophone.XMLDataReader.DescriptorIteration;
import ru.curs.xylophone.XMLDataReader.DescriptorOutput;

public class TestReader {
	private InputStream descrStream;
	private InputStream dataStream;

	@Rule
	public ExpectedException expectedException = ExpectedException.none();

	@After
	public void TearDown() {
		try {
			if (descrStream != null)
				descrStream.close();
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		try {
			if (dataStream != null)
				dataStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	@Test
	public void testCompareNames() {
		HashMap<String, String> a = new HashMap<>();
		a.put("a1", "v1");
		a.put("a2", "v2");
		assertFalse(XMLDataReader.compareNames("bb", "aa", a));
		assertTrue(XMLDataReader.compareNames("bb", "bb", a));
		assertTrue(XMLDataReader.compareNames("*", "aa", a));
		assertTrue(XMLDataReader.compareNames("*", "bb", a));
		assertFalse(XMLDataReader.compareNames("aa", "*", a));

		assertFalse(XMLDataReader.compareNames("bb[@a1=\"v1\"]", "aa", a));
		assertTrue(XMLDataReader.compareNames("aa[@a1=\"v1\"]", "aa", a));
		assertTrue(XMLDataReader.compareNames("aa[@a2=\"v2\"]", "aa", a));
		assertFalse(XMLDataReader.compareNames("aa[@a1=\"v2\"]", "aa", a));
		assertFalse(XMLDataReader.compareNames("aa[@a3=\"v2\"]", "aa", a));

		assertFalse(XMLDataReader.compareNames("bb[@a1='v1']", "aa", a));
		assertTrue(XMLDataReader.compareNames("aa[@a1='v1']", "aa", a));
		assertTrue(XMLDataReader.compareNames("aa[@a2='v2']", "aa", a));
		assertFalse(XMLDataReader.compareNames("aa[@a1='v2']", "aa", a));
		assertFalse(XMLDataReader.compareNames("aa[@a3='v2']", "aa", a));

		assertFalse(XMLDataReader.compareNames("aa[@a1='v1\"]", "aa", a));
		assertFalse(XMLDataReader.compareNames("aa[@a2='v2\"]", "aa", a));
		assertFalse(XMLDataReader.compareNames("aa[@a1=\"v1']", "aa", a));
		assertFalse(XMLDataReader.compareNames("aa[@a2=\"v2']", "aa", a));
	}

	@Test
	public void testParseDescriptor() throws XML2SpreadSheetError {
		descrStream = TestReader.class
				.getResourceAsStream("testdescriptor.xml");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		XMLDataReader reader = XMLDataReader.createReader(dataStream,
				descrStream, false, null);
		DescriptorElement d = reader.getDescriptor();

		assertEquals(2, d.getSubelements().size());
		assertEquals("report", d.getElementName());
		DescriptorIteration i = (DescriptorIteration) d.getSubelements().get(0);
		assertEquals(0, i.getIndex());

		assertFalse(i.isHorizontal());
		DescriptorElement de = i.getElements().get(0);
		i = (DescriptorIteration) de.getSubelements().get(1);

		de = i.getElements().get(0);

		assertEquals("line", de.getElementName());
		assertFalse(((DescriptorOutput) de.getSubelements().get(0))
				.getPageBreak());

		de = i.getElements().get(1);
		assertEquals("group", de.getElementName());
		assertTrue(((DescriptorOutput) de.getSubelements().get(0))
				.getPageBreak());

		i = (DescriptorIteration) d.getSubelements().get(1);
		assertEquals(-1, i.getIndex());
		assertFalse(i.isHorizontal());

		assertEquals(1, i.getElements().size());
		d = i.getElements().get(0);
		assertEquals("sheet", d.getElementName());
		assertEquals(4, d.getSubelements().size());
		DescriptorOutput o = (DescriptorOutput) d.getSubelements().get(0);
		assertEquals("~{@name}", o.getWorksheet());
		o = (DescriptorOutput) d.getSubelements().get(1);
		assertNull(o.getWorksheet());
		i = (DescriptorIteration) d.getSubelements().get(2);
		assertEquals(-1, i.getIndex());
		assertTrue(i.isHorizontal());

	}

	@Test
	public void testDOMReader1() throws XML2SpreadSheetError {
		descrStream = TestReader.class
				.getResourceAsStream("testdescriptor.xml");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		DummyWriter w = new DummyWriter();
		// Проверяем последовательность генерируемых ридером команд
		XMLDataReader reader = XMLDataReader.createReader(dataStream,
				descrStream, false, w);
		reader.process();
		assertEquals(
				"Q{TCQ{CCbQ{CC}C}}Q{TCQh{CCC}Q{CQh{CCC}CQh{CCC}CQh{CCC}}TCQh{}Q{}}F",
				w.getLog().toString());
	}

	@Test
	public void testDOMReader2() throws XML2SpreadSheetError {
		descrStream = TestReader.class
				.getResourceAsStream("testdescriptor2.xml");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		DummyWriter w = new DummyWriter();
		// Проверяем последовательность генерируемых ридером команд
		XMLDataReader reader = XMLDataReader.createReader(dataStream,
				descrStream, false, w);
		reader.process();
		assertEquals("Q{TCQ{CCQ{CC}C}}Q{TCQh{CCC}Q{CQh{CCC}}TCQh{}Q{}}F", w
				.getLog().toString());
	}

	@Test
	public void testSAXReader1() throws XML2SpreadSheetError {
		descrStream = TestReader.class
				.getResourceAsStream("testdescriptor.xml");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		DummyWriter w = new DummyWriter();

		XMLDataReader reader = XMLDataReader.createReader(dataStream,
				descrStream, true, w);
		// Проверяем, что на некорректных данных выскакивает корректное
		// сообщение об ошибке
		boolean itHappened = false;
		try {
			reader.process();
		} catch (XML2SpreadSheetError e) {
			itHappened = true;
			assertTrue(e.getMessage().contains(
					"only one iteration element is allowed"));
		}
		assertTrue(itHappened);

	}

	@Test
	public void testSAXReader2() throws XML2SpreadSheetError {
		descrStream = TestReader.class
				.getResourceAsStream("testsaxdescriptor.xml");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		DummyWriter w = new DummyWriter();
		// Проверяем последовательность генерируемых ридером команд
		XMLDataReader reader = XMLDataReader.createReader(dataStream,
				descrStream, false, w);
		reader.process();
		assertEquals("Q{TCQ{CCQ{CC}C}TQ{CQh{CCC}CQh{CCC}CQh{CCC}}TQ{}}F", w
				.getLog().toString());
	}

	@Test
	public void testSAXReader3() throws XML2SpreadSheetError {
		descrStream = TestReader.class
				.getResourceAsStream("testsaxdescriptor2.xml");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		DummyWriter w = new DummyWriter();
		// Проверяем последовательность генерируемых ридером команд
		XMLDataReader reader = XMLDataReader.createReader(dataStream,
				descrStream, false, w);
		reader.process();
		assertEquals("Q{TCQ{CCQ{CC}C}TQ{CQh{CCC}}TQ{}}F", w.getLog().toString());
	}

	@Test
	public void testParsingDescriptorWithElementInsideElementShouldFail() throws XML2SpreadSheetError {
		descrStream = TestReader.class
				.getResourceAsStream("test_descriptor_with_element_inside_element.xml");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		expectedException.expect(XML2SpreadSheetError.class);
		expectedException.expectMessage("Tag <element> is not allowed inside <element>. Error inside element with name titlepage.");

		DummyWriter w = new DummyWriter();
		// When reader is created, exception is thrown because of not correct sequence of tags
		XMLDataReader.createReader(dataStream,
				descrStream, false, w);
	}

	@Test
	public void testParsingDescriptorWithIterationInsideIterationShouldFail() throws XML2SpreadSheetError {
		descrStream = TestReader.class
				.getResourceAsStream("test_descriptor_with_iteration_inside_iteration.xml");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		expectedException.expect(XML2SpreadSheetError.class);
		expectedException.expectMessage("Tag <iteration> is not allowed inside <iteration>. " +
				"Error inside element with name titlepage.");

		DummyWriter w = new DummyWriter();
		// When reader is created, exception is thrown because of not correct sequence of tags
		XMLDataReader.createReader(dataStream,
				descrStream, false, w);
	}

	@Test
	public void testParsingDescriptorWithOutputInsideIterationShouldFail() throws XML2SpreadSheetError {
		descrStream = TestReader.class
				.getResourceAsStream("test_descriptor_with_output_inside_iteration.xml");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		expectedException.expect(XML2SpreadSheetError.class);
		expectedException.expectMessage("Tag <output> is not allowed inside <iteration>. " +
				"Error inside element with name titlepage.");

		DummyWriter w = new DummyWriter();
		// When reader is created, exception is thrown because of not correct sequence of tags
		XMLDataReader.createReader(dataStream,
				descrStream, false, w);
	}
}

class DummyWriter extends ReportWriter {

	private final String[] sheetNames = { "Титульный", "Раздел А", "Раздел Б" };
	private int i;
	private final StringBuilder log = new StringBuilder();

	@Override
	public void sheet(String sheetName, String sourceSheet,
			int startRepeatingColumn, int endRepeatingColumn,
			int startRepeatingRow, int endRepeatingRow) {
		// sheeT
		log.append("T");
		assertEquals(sheetNames[i], sheetName);
		i++;

		assertEquals(-1, startRepeatingColumn);
		assertEquals(-1, endRepeatingColumn);
		assertEquals(-1, startRepeatingRow);
		assertEquals(-1, endRepeatingColumn);
	}

	@Override
	public void startSequence(boolean horizontal) {
		// seQuence
		if (horizontal)
			log.append("Qh{");
		else
			log.append("Q{");
	}

	@Override
	public void endSequence(int merge, String regionName) {
		log.append("}");
	}

	@Override
	public void section(XMLContext context, String sourceSheet,
			RangeAddress range, boolean pagebreak) {
		// seCtion
		log.append("C");
		if (pagebreak)
			log.append("b");
	}

	StringBuilder getLog() {
		return log;
	}

	@Override
	void newSheet(String sheetName, String sourceSheetint,
			int startRepeatingColumn, int endRepeatingColumn,
			int startRepeatingRow, int endRepeatingRow) {
	}

	@Override
	void putSection(XMLContext context, CellAddress growthPoint2,
			String sourceSheet, RangeAddress range) {
	}

	@Override
	public void flush() {
		// Также проверяем, что последним всегда вызывается метод flush.
		log.append("F");
	}

	@Override
	void mergeUp(CellAddress a1, CellAddress a2) {
		log.append("merge");
	}

	@Override
	void addNamedRegion(String name, CellAddress a1, CellAddress a2) {
		log.append("addNamedRegion");
	}

	@Override
	void putRowBreak(int rowNumber) {
		log.append(String.format("[rowbreak%d]", rowNumber));

	}

	@Override
	void putColBreak(int colNumber) {
		log.append(String.format("[colbreak%d]", colNumber));

	}

}
