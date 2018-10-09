package ru.curs.xylophone;

import org.apache.poi.ss.usermodel.Cell;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Wrapper for cell with style properties.
 */
public final class DynamicCellWithStyle {
    private static final String KEY_REGEX = "([a-zA-Z][a-zA-Z0-9_\\-]*[a-zA-Z-0-9])";
    private static final String VALUE_REGEX = "\"(([^\"]|\"\")*)\"";
    private static final String PROPERTY_REGEX = KEY_REGEX + "\\s*:\\s*" + VALUE_REGEX;

    private static final Pattern PROPERTY_PATTERN = Pattern.compile(PROPERTY_REGEX);
    private static final Pattern PROPERTIES_PATTERN =
            Pattern.compile("\\|\\s*((" + PROPERTY_REGEX + "\\s*;?\\s*)+)$");

    private Cell cell;
    private Map<String, String> properties;
    private String value;

    private DynamicCellWithStyle(Cell cell, Map<String, String> properties, String value) {
        this.cell = cell;
        this.properties = properties;
        this.value = value;
    }

    /**
     * Create DynamicCell with Style Map.
     *
     * @param cell   cell of book
     * @param record string value that we want to write to cell
     * @return DynamicCell where can be map of properties
     */
    public static DynamicCellWithStyle defineCellStyle(Cell cell, String record) {
        Matcher matcher = PROPERTIES_PATTERN.matcher(record);
        if (matcher.find()) {
            String value = record.substring(0, matcher.start());
            Map<String, String> properties = parseProperties(matcher.group(1));
            return new DynamicCellWithStyle(cell, properties, value);
        } else {
            return new DynamicCellWithStyle(cell, Collections.emptyMap(), record);
        }
    }

    /**
     * Getter for cell.
     *
     * @return cell
     */
    public Cell getCell() {
        return cell;
    }

    /**
     * Getter for properties.
     *
     * @return properties map
     */
    public Map<String, String> getProperties() {
        return properties;
    }

    /**
     * Getter for record of cell.
     *
     * @return record of cell
     */
    public String getValue() {
        return value;
    }

    /**
     * Method that speak if style properties for this cell.
     *
     * @return true if map with styles is not empty, otherwise - false
     */
    public boolean isStylesPresent() {
        return !properties.isEmpty();
    }

    /**
     * Parse each element (key: "value") of array and put pair into Map.
     *
     * @param stylePairs style pair - key-value pair for style.
     * @return Map where key is property name and value is property value
     */
    private static Map<String, String> parseProperties(String stylePairs) {
        Map<String, String> properties = new HashMap<>();
        Matcher matcher = PROPERTY_PATTERN.matcher(stylePairs);
        while (matcher.find()) {
            String key = matcher.group(1);
            String value = matcher.group(2).replaceAll("\"\"", "\"");
            properties.put(key, value);
        }
        return properties;
    }
}
