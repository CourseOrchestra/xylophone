package ru.curs.xylophone;

import org.apache.poi.ss.util.CellRangeAddress;
import org.approvaltests.Approvals;
import org.junit.After;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;
import ru.curs.xylophone.descriptor.DescriptorElement;
import ru.curs.xylophone.descriptor.DescriptorIteration;
import ru.curs.xylophone.descriptor.DescriptorOutput;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.stream.Stream;

import static org.junit.Assert.*;

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
	public void testParseDescriptor() throws XylophoneError {
		descrStream = TestReader.class
				.getResourceAsStream("testdescriptor.json");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		XMLDataReader reader = XMLDataReader.createReader(dataStream,
				descrStream, false,null);
		DescriptorElement d = reader.getDescriptor();

		assertEquals(2, d.getSubElements().size());
		assertEquals("report", d.getName());
		DescriptorIteration i = (DescriptorIteration) d.getSubElements().get(0);
		assertEquals(0, i.getIndex());

		assertFalse(i.isHorizontal());
		DescriptorElement de = i.getElements().get(0);
		i = (DescriptorIteration) de.getSubElements().get(1);

		de = i.getElements().get(0);

		assertEquals("line", de.getName());
		assertFalse(((DescriptorOutput) de.getSubElements().get(0))
				.getPageBreak());

		de = i.getElements().get(1);
		assertEquals("group", de.getName());
		assertTrue(((DescriptorOutput) de.getSubElements().get(0))
				.getPageBreak());

		i = (DescriptorIteration) d.getSubElements().get(1);
		assertEquals(-1, i.getIndex());
		assertFalse(i.isHorizontal());

		assertEquals(1, i.getElements().size());
		d = i.getElements().get(0);
		assertEquals("sheet", d.getName());
		assertEquals(4, d.getSubElements().size());
		DescriptorOutput o = (DescriptorOutput) d.getSubElements().get(0);
		assertEquals("~{@name}", o.getWorksheet());
		o = (DescriptorOutput) d.getSubElements().get(1);
		assertNull(o.getWorksheet());
		i = (DescriptorIteration) d.getSubElements().get(2);
		assertEquals(-1, i.getIndex());
		assertTrue(i.isHorizontal());

	}

	@Test
	public void testDOMReader1() throws XylophoneError {
		descrStream = TestReader.class
				.getResourceAsStream("testdescriptor.json");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		DummyWriter w = new DummyWriter();
		// Проверяем последовательность генерируемых ридером команд
		XMLDataReader reader = XMLDataReader.createReader(dataStream,
				descrStream, false, w);
		reader.process();
		Approvals.verify(w.getLog().toString());
	}

	@Test
	public void testDOMReader2() throws XylophoneError {
		descrStream = TestReader.class
				.getResourceAsStream("testdescriptor2.json");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		DummyWriter w = new DummyWriter();
		// Проверяем последовательность генерируемых ридером команд
		XMLDataReader reader = XMLDataReader.createReader(dataStream,
				descrStream, false, w);
		reader.process();
		Approvals.verify(w.getLog().toString());
	}

	@Test
	public void testSAXReader1() throws XylophoneError {
		descrStream = TestReader.class
				.getResourceAsStream("testdescriptor.json");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		DummyWriter w = new DummyWriter();

		XMLDataReader reader = XMLDataReader.createReader(dataStream,
				descrStream, true, w);
		// Проверяем, что на некорректных данных выскакивает корректное
		// сообщение об ошибке
		boolean itHappened = false;
		try {
			reader.process();
		} catch (XylophoneError e) {
			itHappened = true;
			assertTrue(e.getMessage().contains(
					"only one iteration element is allowed"));
		}
		assertTrue(itHappened);

	}

	@Test
	public void testSAXReader2() throws XylophoneError {
		descrStream = TestReader.class
				.getResourceAsStream("testsaxdescriptor.json");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		DummyWriter w = new DummyWriter();
		// Проверяем последовательность генерируемых ридером команд
		XMLDataReader reader = XMLDataReader.createReader(dataStream,
				descrStream, true, w);
		reader.process();
		Approvals.verify(w.getLog().toString());
	}

	@Test
	public void testSAXReader3() throws XylophoneError {
		descrStream = TestReader.class
				.getResourceAsStream("testsaxdescriptor2.json");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		DummyWriter w = new DummyWriter();
		// Проверяем последовательность генерируемых ридером команд
		XMLDataReader reader = XMLDataReader.createReader(dataStream,
				descrStream, true, w);
		reader.process();
		Approvals.verify(w.getLog().toString());
	}

	@Test
	public void testParsingDescriptorWithElementInsideElementShouldFail() throws XylophoneError {
		descrStream = TestReader.class
				.getResourceAsStream("test_descriptor_with_element_inside_element.json");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		expectedException.expect(XylophoneError.class);
		expectedException.expectMessage("Error while processing json descriptor: " +
				"Unrecognized field \"element\" (class ru.curs.xylophone.descriptor.DescriptorElement), not marked as ignorable (2 known properties: \"name\", \"output-steps\"])");

		DummyWriter w = new DummyWriter();
		// When reader is created, exception is thrown because of not correct sequence of tags
		XMLDataReader.createReader(dataStream,
				descrStream, false, w);
	}

	@Test
	public void testParsingDescriptorWithIterationInsideIterationShouldFail() throws XylophoneError {
		descrStream = TestReader.class
				.getResourceAsStream("test_descriptor_with_iteration_inside_iteration.json");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

		expectedException.expect(XylophoneError.class);
    expectedException.expectMessage("Error while processing json descriptor: " +
                "Cannot deserialize instance of `ru.curs.xylophone.descriptor.DescriptorIteration` out of START_ARRAY token");

		DummyWriter w = new DummyWriter();
		// When reader is created, exception is thrown because of not correct sequence of tags
		XMLDataReader.createReader(dataStream,
				descrStream, false, w);
	}

	@Test
	public void testParsingDescriptorWithOutputInsideIterationShouldFail() throws XylophoneError {
		descrStream = TestReader.class
				.getResourceAsStream("test_descriptor_with_output_inside_iteration.json");
		dataStream = TestReader.class.getResourceAsStream("testdata.xml");

  	expectedException.expect(XylophoneError.class);
    expectedException.expectMessage("Error while processing json descriptor: " +
                "Cannot deserialize instance of `ru.curs.xylophone.descriptor.DescriptorIteration` out of START_ARRAY token");

		DummyWriter w = new DummyWriter();
		// When reader is created, exception is thrown because of not correct sequence of tags
		XMLDataReader.createReader(dataStream,
				descrStream, false, w);
	}
}

class DummyWriter extends ReportWriter {

	private final String[] sheetNames = { "Титульный", "Раздел А", "Раздел Б" };
	private int sheetNo;
	private int offsetSize = 0;
	private final StringBuilder log = new StringBuilder();

	public void writeLine(String line) {
		for (int i = 0; i < offsetSize; i++)
			log.append("\t");
		log.append(line).append("\n");
	}

	public void startLogSection(String line) {
		writeLine(line);
		offsetSize += 1;
	}

	public void endLogSection(String line) {
		offsetSize -= 1;
		writeLine(line);
	}

	@Override
	public void sheet(String sheetName, String sourceSheet,
					  int startRepeatingColumn, int endRepeatingColumn,
					  int startRepeatingRow, int endRepeatingRow) {
		startLogSection("Creating sheet " + sheetName);
		assertEquals(sheetNames[sheetNo], sheetName);
		sheetNo++;

		assertEquals(-1, startRepeatingColumn);
		assertEquals(-1, endRepeatingColumn);
		assertEquals(-1, startRepeatingRow);
		assertEquals(-1, endRepeatingColumn);
	}

	@Override
	public void startSequence(boolean horizontal) {
		if (horizontal)
			startLogSection("Starting horizontal sequence");
		else
			startLogSection("Starting vertical sequence");
	}

	@Override
	public void endSequence(int merge, String regionName) {
		endLogSection("Finalizing sequence");
	}

	@Override
	public void section(XMLContext context, String sourceSheet,
						RangeAddress range, boolean pagebreak) {
		String message = "Section ";
		if (sourceSheet != null && !sourceSheet.equals(""))
			message += "from sheet " + sourceSheet + " ";
		if (range != null) {
			message += "(range " + range.getAddress() + ")";
		}
		if (pagebreak)
			message += " with page break";

		writeLine(message);
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
		writeLine("Flush");
	}

	@Override
	void mergeUp(CellAddress a1, CellAddress a2) {
		writeLine("Merging cells " + a1.getAddress() + " and " + a2.getAddress() + " up ");
	}

	@Override
	void addNamedRegion(String name, CellAddress a1, CellAddress a2) {
		writeLine("Adding named region (" + a1.getAddress() + ":" + a2.getAddress() + ")");
	}

	@Override
	void putRowBreak(int rowNumber) {
		writeLine("Putting row break at row number" + rowNumber);
	}

	@Override
	void putColBreak(int colNumber) {
		writeLine("Putting column break at column number" + colNumber);

	}

	@Override
	void applyMergedRegions(Stream<CellRangeAddress> mergedRegions){
	}

}

