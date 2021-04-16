package ru.curs.xylophone;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class TestCellRangeAddress {

	@Test
	public void testCellAddress() {
		CellAddress ca = new CellAddress("D12");
		assertEquals(4, ca.getCol());
		assertEquals(12, ca.getRow());
		assertEquals("D12", ca.getAddress());

		ca = new CellAddress("AB11");
		assertEquals(28, ca.getCol());
		assertEquals(11, ca.getRow());
		assertEquals("AB11", ca.getAddress());

		assertEquals(new CellAddress("ZA111"), new CellAddress("ZA111"));
	}

	@Test
	public void testCellAddress2() {
		CellAddress ca = new CellAddress("Z1");
		assertEquals(26, ca.getCol());
		assertEquals("Z1", ca.getAddress());
		ca = new CellAddress("AZ21");
		assertEquals(52, ca.getCol());
		assertEquals("AZ21", ca.getAddress());
		ca = new CellAddress("BX21");
		assertEquals(76, ca.getCol());
		assertEquals("BX21", ca.getAddress());
	}

	@Test
	public void testRangeAddress1() throws XylophoneError {
		RangeAddress ra = new RangeAddress("D11:G36");
		assertEquals(new CellAddress("D11"), ra.topLeft());
		assertEquals(new CellAddress("G36"), ra.bottomRight());
		assertEquals(4, ra.left());
		assertEquals(7, ra.right());
		assertEquals(11, ra.top());
		assertEquals(36, ra.bottom());
	}

	@Test
	public void testRangeAddress2() throws XylophoneError {
		// Автонормализация диапазона
		RangeAddress ra = new RangeAddress("G36:D11");
		assertEquals(new CellAddress("D11"), ra.topLeft());
		assertEquals(new CellAddress("G36"), ra.bottomRight());
		assertEquals(4, ra.left());
		assertEquals(7, ra.right());
		assertEquals(11, ra.top());
		assertEquals(36, ra.bottom());

	}

	@Test
	public void testRangeAddress3() throws XylophoneError {
		// Диапазон из одной ячейки
		RangeAddress ra = new RangeAddress("E7");
		assertEquals(new CellAddress("E7"), ra.topLeft());
		assertEquals(new CellAddress("E7"), ra.bottomRight());
		assertEquals(5, ra.left());
		assertEquals(5, ra.right());
		assertEquals(7, ra.top());
		assertEquals(7, ra.bottom());
	}

}
