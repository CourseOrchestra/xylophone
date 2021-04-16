package ru.curs.xylophone.descriptor;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonGetter;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;
import ru.curs.xylophone.RangeAddress;
import ru.curs.xylophone.XML2SpreadSheetError;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

@JsonIgnoreProperties(value = {
        "startRepeatingColumn",
        "endRepeatingColumn",
        "startRepeatingRow",
        "endRepeatingRow",
        "pageBreak"})
public final class DescriptorOutput extends DescriptorOutputBase {
    private static final Pattern RANGE = Pattern
            .compile("(-?[0-9]+):(-?[0-9]+)");

    private final String worksheet;
    private final RangeAddress range;
    private final String sourceSheet;
    private final int startRepeatingColumn;
    private final int endRepeatingColumn;
    private final int startRepeatingRow;
    private final int endRepeatingRow;
    private final boolean pageBreak;

    @JsonCreator(mode = JsonCreator.Mode.PROPERTIES)
    DescriptorOutput(
            @JsonProperty("worksheet") String worksheet,
            @JsonProperty("range") String range,
            @JsonProperty("sourcesheet") String sourceSheet,
            @JsonProperty("repeatingcols") String repeatingCols,
            @JsonProperty("repeatingraws") String repeatingRows,
            @JsonProperty("pagebreak") String pageBreak) throws XML2SpreadSheetError {
        this(
                worksheet,
                range == null ? null : new RangeAddress(range),
                sourceSheet,
                repeatingCols,
                repeatingRows,
                pageBreak != null &&
                        (pageBreak.equalsIgnoreCase("true") ||
                                !pageBreak.equalsIgnoreCase("0"))
        );
    }

    public DescriptorOutput(String worksheet, RangeAddress range,
                            String sourceSheet, String repeatingCols, String repeatingRows,
                            Boolean pageBreak) throws XML2SpreadSheetError {
        this.worksheet = worksheet;
        this.range = range;
        this.sourceSheet = sourceSheet;
        this.pageBreak = pageBreak;
        Matcher m1 = RANGE.matcher(repeatingCols == null ? "-1:-1"
                : repeatingCols);
        Matcher m2 = RANGE.matcher(repeatingRows == null ? "-1:-1"
                : repeatingRows);
        if (m1.matches() && m2.matches()) {
            this.startRepeatingColumn = Integer.parseInt(m1.group(1));
            this.endRepeatingColumn = Integer.parseInt(m1.group(2));
            this.startRepeatingRow = Integer.parseInt(m2.group(1));
            this.endRepeatingRow = Integer.parseInt(m2.group(2));
        } else {
            throw new XML2SpreadSheetError(String.format(
                    "Invalid col/row range %s %s", repeatingCols,
                    repeatingRows));
        }
    }

    @JsonGetter("worksheet")
    public String getWorksheet() {
        return worksheet;
    }

    @JsonGetter("sourcesheet")
    public String getSourceSheet() {
        return sourceSheet;
    }

    @JsonGetter("range")
    public String getRangeString() {
        if (range == null) return null;

        String bottomRight = range.bottomRight().getAddress();
        String topLeft = range.topLeft().getAddress();
        if (bottomRight.equals(topLeft))
            return bottomRight;
        return range.getAddress();
    }

    @JsonGetter("repeatingcols")
    public String getRepeatingCols() {
        if (startRepeatingColumn == -1 && endRepeatingColumn == -1)
            return null;
        return startRepeatingColumn + ":" + endRepeatingColumn;
    }

    @JsonGetter("repeatingrows")
    public String getRepeatingRows() {
        if (startRepeatingRow == -1 && endRepeatingRow == -1)
            return null;
        return startRepeatingRow + ":" + endRepeatingRow;
    }

    @JsonGetter("pagebreak")
    public String getPageBreakString() {
        return pageBreak ? "true" : null;
    }

    public boolean getPageBreak() {
        return pageBreak;
    }

    public RangeAddress getRange() {
        return range;
    }

    public int getStartRepeatingColumn() {
        return startRepeatingColumn;
    }

    public int getEndRepeatingColumn() {
        return endRepeatingColumn;
    }

    public int getStartRepeatingRow() {
        return startRepeatingRow;
    }

    public int getEndRepeatingRow() {
        return endRepeatingRow;
    }
}
