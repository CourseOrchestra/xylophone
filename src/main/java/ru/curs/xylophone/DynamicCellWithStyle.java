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
    private static final Pattern PROPERTIES_PATTERN =
            Pattern.compile("\\|\\|\\s*([a-zA-Z][a-zA-Z0-9_\\-]*[a-zA-Z-0-9]\\s*:\\s*\"([^\"]|\"\")*\"\\s*;?\\s*)+$");
    private static final Pattern PROPERTY_PATTERN =
            Pattern.compile("([a-zA-Z][a-zA-Z0-9_\\-]*[a-zA-Z-0-9]\\s*:\\s*\"(([^\"]|\"\")*)\")");
    private static final Pattern KEY_PATTERN =
            Pattern.compile("[a-zA-Z][a-zA-Z0-9_\\-]*[a-zA-Z-0-9]");
    private static final Pattern VALUE_PATTERN =
            Pattern.compile("\"(([^\"]|\"\")*)\"");

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
     * @param cell cell of book
     * @param record string value that we want to write to cell
     * @return DynamicCell where can be map of properties
     */
    public static DynamicCellWithStyle defineCellStyle(Cell cell, String record) {
        Matcher matcher = PROPERTIES_PATTERN.matcher(record);

        if (!matcher.find()) {
            return new DynamicCellWithStyle(cell, Collections.emptyMap(), record);
        }

        int start = matcher.start();
        int end = matcher.end();

        String propertyString = record.substring(start, end);

        Map<String, String> properties = parseProperties(propertyString);

        return new DynamicCellWithStyle(cell, properties, record.substring(0, start));
    }

    /**
     * Getter for cell.
     * @return cell
     */
    public Cell getCell() {
        return cell;
    }

    /**
     * Getter for properties.
     * @return properties map
     */
    public Map<String, String> getProperties() {
        return properties;
    }

    /**
     * Getter for record of cell.
     * @return record of cell
     */
    public String getValue() {
        return value;
    }

    /**
     * Method that speak if style properties for this cell.
     * @return true if map with styles is not empty, otherwise - false
     */
    public boolean isStylesPresent() {
        return !properties.isEmpty();
    }

    /**
     * Parse each element (key: "value") of array and put pair into Map.
     * @param stylePairs style pair - key-value pair for style.
     * @return Map where key is property name and value is property value
     */
    private static Map<String, String> parseProperties(String stylePairs) {
        Map<String, String> properties = new HashMap<>();

        Matcher matcher = PROPERTY_PATTERN.matcher(stylePairs);

        while (matcher.find()) {
            int start = matcher.start();
            int end = matcher.end();

            String keyValue = stylePairs.substring(start, end);
            String[] keyValuePair = parseKeyValue(keyValue);

            properties.put(keyValuePair[0], keyValuePair[1]);
        }

        return properties;
    }

    /**
     * Parsing keyValue string to key-value array, where first elem - key, second - value.
     * @param keyValue string representation of key-value param
     * @return array with the first element which is key, and the second is value
     */
    private static String[] parseKeyValue(String keyValue) {
        Matcher keyMatcher = KEY_PATTERN.matcher(keyValue);
        Matcher valueMatcher = VALUE_PATTERN.matcher(keyValue);

        String[] keyValuePair = new String[2];

        if (keyMatcher.find() && valueMatcher.find()) {
            int startOfKey = keyMatcher.start();
            int endOfKey = keyMatcher.end();

            int startOfValue = valueMatcher.start();
            int endOfValue = valueMatcher.end();

            keyValuePair[0] = keyValue.substring(startOfKey, endOfKey);
            //For remove first and last quotes; also merge double quotes
            keyValuePair[1] = keyValue.substring(startOfValue + 1, endOfValue - 1)
                    .replaceAll("\"\"", "\"");
        } else {
            throw new RuntimeException(
                    String.format("Find properties part of string, but one of properties %s is not valid", keyValue));
        }

        return keyValuePair;
    }
}
