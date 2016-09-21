package ru.curs.xylophone;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class TestXMLContext {
	@Test
	public void testCalc() {
		DummyContext dc = new DummyContext();
		assertEquals("YYY", dc.calc("~{bbb}"));
		assertEquals("aaaXXXbbbYYYZZZ", dc.calc("aaa~{aaa}bbb~{bbb}~{foo}"));
	}
}

class DummyContext extends XMLContext {
	@Override
	String getXPathValue(String xpath) {
		if ("aaa".equals(xpath))
			return "XXX";
		else if ("bbb".equals(xpath))
			return "YYY";
		else
			return "ZZZ";
	}

}
