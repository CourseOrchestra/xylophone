package ru.curs.xylophone;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.junit.Test;

public class TestFO {
	@Test
	public void test1() throws Exception {
		InputStream is = TestFO.class.getResourceAsStream("template.xls");
		OutputStream os = new ByteArrayOutputStream();
		// os = new FileOutputStream("d:/temp/xls.pdf");
		try {
			Excel2Print e2p = new Excel2Print(is);
			/*
			 * e2p.setFopConfig( new File(
			 * "D:/workspaces/workspace/xylophone/src/test/resources/ru/curs/xylophone/fop.xconf"
			 * ));
			 */
			e2p.toFO(os);

			// e2p.toPDF(os);
		} finally {
			is.close();
			os.close();
		}
	}
}
