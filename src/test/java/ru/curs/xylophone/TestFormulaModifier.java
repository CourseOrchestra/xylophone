package ru.curs.xylophone;

import org.junit.Test;

import static org.junit.Assert.assertEquals;
import static ru.curs.xylophone.FormulaModifier.modifyFormula;

public class TestFormulaModifier {
	@Test
	public void test1() {
		String f = "SUM(AA501:BB23)+E15";
		assertEquals("SUM(AB502:BC24)+F16", modifyFormula(f, 1, 1));
		assertEquals("SUM(Z500:BA22)+D14", modifyFormula(f, -1, -1));
	}
	
	@Test
	public void test2() {
		String f = "234 + G4 + F4";
		assertEquals("234 + G5 + F5", modifyFormula(f, 0, 1));
		assertEquals("234 + I4 + H4", modifyFormula(f, 2, 0));
	}
	
	@Test
	public void test3() {
		String f = "=ATAN2(B3;B5)+LOG10(B3)";
		assertEquals("=ATAN2(B4;B6)+LOG10(B4)", modifyFormula(f, 0, 1));
		assertEquals("=ATAN2(D3;D5)+LOG10(D3)", modifyFormula(f, 2, 0));
	}
	
	
}
