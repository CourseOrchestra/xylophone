package ru.curs.xylophone;

import org.junit.Test;

import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

public class DynamicCellWithStyleUT {

    @Test
    public void test01() {
        String testData = "Some Text {@foo} |backgroundcolor:\"#ffffff\";";

        DynamicCellWithStyle cellStyle = DynamicCellWithStyle.defineCellStyle(null, testData);

        assertTrue(cellStyle.isStylesPresent());

        Map<String, String> expected = new HashMap<>();
        expected.put("backgroundcolor", "#ffffff");

        Map<String, String> actual = cellStyle.getProperties();

        assertTrue(actual.equals(expected));
        assertTrue(cellStyle.getValue().equals("Some Text {@foo} "));
    }

    @Test
    public void test02() {
        String testData = "Some Text ~{@fBar} bar|color:\"#353833\"";

        DynamicCellWithStyle cellStyle = DynamicCellWithStyle.defineCellStyle(null, testData);

        assertTrue(cellStyle.isStylesPresent());

        Map<String, String> expected = new HashMap<>();
        expected.put("color", "#353833");

        Map<String, String> actual = cellStyle.getProperties();

        assertTrue(actual.equals(expected));
        assertTrue(cellStyle.getValue().equals("Some Text ~{@fBar} bar"));
    }

    @Test
    public void test03() {
        String testData = "|    fontfamily:\"'DejaVu Sans', Arial, Helvetica, sans-serif\";";

        DynamicCellWithStyle cellStyle = DynamicCellWithStyle.defineCellStyle(null, testData);

        assertTrue(cellStyle.isStylesPresent());

        Map<String, String> expected = new HashMap<>();
        expected.put("fontfamily", "'DejaVu Sans', Arial, Helvetica, sans-serif");

        Map<String, String> actual = cellStyle.getProperties();

        assertTrue(actual.equals(expected));
        assertTrue(cellStyle.getValue().equals(""));
    }

    @Test
    public void test04() {
        String testData = "foo text|    fontsize:\"14px\";";

        DynamicCellWithStyle cellStyle = DynamicCellWithStyle.defineCellStyle(null, testData);

        assertTrue(cellStyle.isStylesPresent());

        Map<String, String> expected = new HashMap<>();
        expected.put("fontsize", "14px");

        Map<String, String> actual = cellStyle.getProperties();

        assertTrue(actual.equals(expected));
        assertTrue(cellStyle.getValue().equals("foo text"));
    }

    @Test
    public void test05() {
        String testData = "|    margin:\"0\"";

        DynamicCellWithStyle cellStyle = DynamicCellWithStyle.defineCellStyle(null, testData);

        assertTrue(cellStyle.isStylesPresent());

        Map<String, String> expected = new HashMap<>();
        expected.put("margin", "0");

        Map<String, String> actual = cellStyle.getProperties();

        assertTrue(actual.equals(expected));
        assertTrue(cellStyle.getValue().equals(""));
    }

    @Test
    public void test06() {
        String testData = "Some Text {@foo} |backgroundcolor:\"#ffffff\";" +
                "color:\"#353833\";" +
                "  fontfamily:\"'DejaVu Sans', Arial, Helvetica, sans-serif\";" +
                "fontsize:\"14px\";" +
                "    margin:\"0\"";

        DynamicCellWithStyle cellStyle = DynamicCellWithStyle.defineCellStyle(null, testData);

        assertTrue(cellStyle.isStylesPresent());

        Map<String, String> expected = new HashMap<>();
        expected.put("backgroundcolor", "#ffffff");
        expected.put("color", "#353833");
        expected.put("fontfamily", "'DejaVu Sans', Arial, Helvetica, sans-serif");
        expected.put("fontsize", "14px");
        expected.put("margin", "0");

        Map<String, String> actual = cellStyle.getProperties();

        assertTrue(actual.equals(expected));
        assertTrue(cellStyle.getValue().equals("Some Text {@foo} "));
    }

    @Test
    public void test07() {
        String testData = "sss asss |color: \"#AABBCC\"; value: \"a&nbsp;b\"; quotedvalue: \"aa\"\"bb\"\"\"";

        DynamicCellWithStyle cellStyle = DynamicCellWithStyle.defineCellStyle(null, testData);

        assertTrue(cellStyle.isStylesPresent());

        Map<String, String> expected = new HashMap<>();
        expected.put("color", "#AABBCC");
        expected.put("value", "a&nbsp;b");
        expected.put("quotedvalue", "aa\"bb\"");

        Map<String, String> actual = cellStyle.getProperties();

        assertTrue(actual.equals(expected));
        assertTrue(cellStyle.getValue().equals("sss asss "));
    }

    @Test
    public void test08() {
        String testData = "aaa | aa; bb";

        DynamicCellWithStyle cellWithStyle = DynamicCellWithStyle.defineCellStyle(null, testData);

        assertFalse(cellWithStyle.isStylesPresent());

        assertTrue(cellWithStyle.getValue().equals(testData));
    }

    @Test
    public void test09() {
        String testData = "aaa abbb |  " +
                "color: \"#AABBCC\"; value: \"a&nbsp;b\"; quotedvalue: \"aa\"\"bb\"\"\"; background-color: \"blue\"";

        DynamicCellWithStyle cellStyle = DynamicCellWithStyle.defineCellStyle(null, testData);

        assertTrue(cellStyle.isStylesPresent());

        Map<String, String> expected = new HashMap<>();
        expected.put("color", "#AABBCC");
        expected.put("value", "a&nbsp;b");
        expected.put("quotedvalue", "aa\"bb\"");
        expected.put("background-color", "blue");

        Map<String, String> actual = cellStyle.getProperties();

        assertTrue(actual.equals(expected));
        assertTrue(cellStyle.getValue().equals("aaa abbb "));
    }

    @Test
    public void test10() {
        String testData = "aaa | bbb  |key: \"value | \"";

        DynamicCellWithStyle cellStyle = DynamicCellWithStyle.defineCellStyle(null, testData);

        assertTrue(cellStyle.isStylesPresent());

        Map<String, String> expected = new HashMap<>();
        expected.put("key", "value | ");

        Map<String, String> actual = cellStyle.getProperties();

        assertTrue(actual.equals(expected));
        assertTrue(cellStyle.getValue().equals("aaa | bbb  "));
    }
}
