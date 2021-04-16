package ru.curs.xylophone;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;

import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertTrue;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

public class MergeRegionContainerUT {
    @Rule
    public ExpectedException expectedException = ExpectedException.none();

    /**
     * Test simple left merge.
     * |(1, 1)| <- |(1,2)|
     * Should be success.
     */
    @Test
    public void testMergeLeftOneShouldBeSuccess() {
        MergeRegionContainer container = MergeRegionContainer.getContainer();

        container.mergeLeft(new CellAddress(2, 1));

        List<CellRangeAddress> actualRanges = new ArrayList<>();
        Sheet sheet = mock(Sheet.class);
        when(sheet.addMergedRegion(any(CellRangeAddress.class))).thenAnswer(i -> {
            CellRangeAddress address = i.getArgument(0);
            actualRanges.add(address);
            return actualRanges.size();
        });

        //        container.apply(sheet);
        container.getMergedRegions().forEach(sheet::addMergedRegion);
        container.clear();

        assertTrue(actualRanges.size() == 1);
        assertTrue(new CellRangeAddress(0, 0, 0, 1).equals(actualRanges.get(0)));
    }

    /**
     * Test simple up merge.
     * |(1, 1)| <- |(2, 1)|
     * Should be success
     */
    @Test
    public void testMergeUpOneShouldBeSuccess() {
        MergeRegionContainer container = MergeRegionContainer.getContainer();

        container.mergeUp(new CellAddress(1, 2));

        List<CellRangeAddress> actualRanges = new ArrayList<>();
        Sheet sheet = mock(Sheet.class);
        when(sheet.addMergedRegion(any(CellRangeAddress.class))).thenAnswer(i -> {
            CellRangeAddress address = i.getArgument(0);
            actualRanges.add(address);
            return actualRanges.size();
        });

        container.getMergedRegions().forEach(sheet::addMergedRegion);
        container.clear();

        assertTrue(actualRanges.size() == 1);
        assertTrue(new CellRangeAddress(0, 1, 0, 0).equals(actualRanges.get(0)));
    }

    /**
     * Test left merge for first column.
     * <- |(3, 1)|
     * Should throw IllegalStateException
     */
    @Test
    public void testMergeLeftOneForFirstColumnShouldFail() {
        expectedException.expect(IllegalArgumentException.class);
        expectedException.expectMessage("Cannot merge cell with address A3. It is out of range for left merge");

        MergeRegionContainer container = MergeRegionContainer.getContainer();

        container.mergeLeft(new CellAddress(1, 3));
    }

    /**
     * Test up merge for first row.
     * |(1, 4)| to up
     * Should throw IllegalStateException
     */
    @Test
    public void testMergeUpOneForFirstRowShoulFail() {
        expectedException.expect(IllegalArgumentException.class);
        expectedException.expectMessage("Cannot merge cell with address D1. It is out of range for up merge");

        MergeRegionContainer container = MergeRegionContainer.getContainer();

        container.mergeUp(new CellAddress(4, 1));
    }

    /**
     * Test left merge for several cells.
     * |(1, 1)| <- |(1, 2)| <- |(1, 3)|
     * Should be success
     */
    @Test
    public void testMergeLeftSeveralCellsShouldBeSuccess() {
        MergeRegionContainer container = MergeRegionContainer.getContainer();

        container.mergeLeft(new CellAddress(2, 1));
        container.mergeLeft(new CellAddress(3, 1));

        List<CellRangeAddress> actualRanges = new ArrayList<>();
        Sheet sheet = mock(Sheet.class);
        when(sheet.addMergedRegion(any(CellRangeAddress.class))).thenAnswer(i -> {
            CellRangeAddress address = i.getArgument(0);
            actualRanges.add(address);
            return actualRanges.size();
        });

        container.getMergedRegions().forEach(sheet::addMergedRegion);
        container.clear();

        assertTrue(actualRanges.size() == 1);
        assertTrue(new CellRangeAddress(0, 0, 0, 2).equals(actualRanges.get(0)));
    }

    /**
     * Test up merge for several cells.
     * |(1, 3)| <- |(2, 3)| <- |(3, 3)|
     * Should be success
     */
    @Test
    public void testMergeUpSeveralCellsShouldBeSuccess() {
        MergeRegionContainer container = MergeRegionContainer.getContainer();

        container.mergeUp(new CellAddress(3, 2));
        container.mergeUp(new CellAddress(3, 3));

        List<CellRangeAddress> actualRanges = new ArrayList<>();
        Sheet sheet = mock(Sheet.class);
        when(sheet.addMergedRegion(any(CellRangeAddress.class))).thenAnswer(i -> {
            CellRangeAddress address = i.getArgument(0);
            actualRanges.add(address);
            return actualRanges.size();
        });

        container.getMergedRegions().forEach(sheet::addMergedRegion);
        container.clear();

        assertTrue(actualRanges.size() == 1);
        assertTrue(new CellRangeAddress(0, 2, 2, 2).equals(actualRanges.get(0)));
    }

    /**
     * Test several left and up merges for several cells.
     * |(1, 3)| <- |(2, 3)| <- |(3, 3)|
     * |(4, 1)| <- |(4, 2)| <- |(4, 3)|
     * |(5, 5)| <- |(5, 6)|
     * |(7, 5)| <- |(8, 5)|
     * Should be success
     */
    @Test
    public void testSeveralMergeUpAndSeveralMergeLeftShouldBeSuccess() {
        MergeRegionContainer container = MergeRegionContainer.getContainer();

        container.mergeUp(new CellAddress(3, 2));
        container.mergeUp(new CellAddress(3, 3));

        container.mergeLeft(new CellAddress(2, 4));
        container.mergeLeft(new CellAddress(3, 4));

        container.mergeLeft(new CellAddress(6, 5));

        container.mergeUp(new CellAddress(5, 8));

        List<CellRangeAddress> actualRanges = new ArrayList<>();
        Sheet sheet = mock(Sheet.class);
        when(sheet.addMergedRegion(any(CellRangeAddress.class))).thenAnswer(i -> {
            CellRangeAddress address = i.getArgument(0);
            actualRanges.add(address);
            return actualRanges.size();
        });

        container.getMergedRegions().forEach(sheet::addMergedRegion);
        container.clear();

        assertTrue(actualRanges.size() == 4);
        assertTrue(new CellRangeAddress(0, 2, 2, 2).equals(actualRanges.get(0)));
        assertTrue(new CellRangeAddress(3, 3, 0, 2).equals(actualRanges.get(1)));
        assertTrue(new CellRangeAddress(4, 4, 4, 5).equals(actualRanges.get(2)));
        assertTrue(new CellRangeAddress(6, 7, 4, 4).equals(actualRanges.get(3)));
    }

    /**
     * Test several left and up merges for several cells.
     * Form is not rectangle --- second merge is ignored.
     * |(1, 3)| <- |(2, 3)| <- |(3, 3)|
     * |(1, 3)| <- |(1, 4)|
     * Should be success
     */
    @Test
    public void testSeveralMergeUpAndSeveralMergeWithNoRectangleFormShouldFail() {
        MergeRegionContainer container = MergeRegionContainer.getContainer();

        container.mergeUp(new CellAddress(3, 2));
        container.mergeUp(new CellAddress(3, 3));

        container.mergeLeft(new CellAddress(4, 1));

        List<CellRangeAddress> actualRanges = new ArrayList<>();
        Sheet sheet = mock(Sheet.class);
        when(sheet.addMergedRegion(any(CellRangeAddress.class))).thenAnswer(i -> {
            CellRangeAddress address = i.getArgument(0);
            actualRanges.add(address);
            return actualRanges.size();
        });

        container.getMergedRegions().forEach(sheet::addMergedRegion);
        container.clear();

        assertTrue(actualRanges.size() == 1);
        assertTrue(new CellRangeAddress(0, 2, 2, 2).equals(actualRanges.get(0)));
    }

    @Test
    public void testAddMergionRegion() {
        MergeRegionContainer container = MergeRegionContainer.getContainer();

        CellRangeAddress rangeAddress = new CellRangeAddress(0, 4, 0, 4);
        container.addMergedRegion(rangeAddress);

        List<CellRangeAddress> actualRanges = new ArrayList<>();
        Sheet sheet = mock(Sheet.class);
        when(sheet.addMergedRegion(any(CellRangeAddress.class))).thenAnswer(i -> {
            CellRangeAddress address = i.getArgument(0);
            actualRanges.add(address);
            return actualRanges.size();
        });

        container.getMergedRegions().forEach(sheet::addMergedRegion);
        container.clear();

        assertTrue(actualRanges.size() == 1);
        assertTrue(new CellRangeAddress(0, 4, 0, 4).equals(actualRanges.get(0)));
    }
}
