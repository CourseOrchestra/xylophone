package ru.curs.xylophone;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Container for merged regions.
 * After process sheet merge cells.
 */
public final class MergeRegionContainer {
    private static MergeRegionContainer container;

    private List<MutableCellRangeAddress> mergedRegions = new ArrayList<>();

    /**
     * Private constructor.
     * It's singleton.
     */
    private MergeRegionContainer() {

    }

    /**
     * Singleton getInstance() method.
     * @return instance
     */
    public static MergeRegionContainer getContainer() {
        if (container == null) {
            container = new MergeRegionContainer();
        }

        return container;
    }

    /**
     * Merge with up cell.
     * @param address address of current cell.
     */
    public void mergeUp(CellAddress address) {
        if (address.getRow() == 1) {
            throw new IllegalArgumentException(
                    String.format("Cannot merge cell with address %s. It is out of range for up merge",
                            address.getAddress()));
        }
        CellRangeAddress res = new CellRangeAddress(address.getRow() - 2,
                address.getRow() - 1, address.getCol() - 1, address.getCol() - 1);
        List<MutableCellRangeAddress> intersectedRegion = findIntersectedRegion(res);

        addMergedRegion(res, (List<MutableCellRangeAddress>) intersectedRegion);
    }

    /**
     * Merge with left cell.
     * @param address address of current cell.
     */
    public void mergeLeft(CellAddress address) {
        if (address.getCol() == 1) {
            throw new IllegalArgumentException(
                    String.format("Cannot merge cell with address %s. It is out of range for left merge",
                            address.getAddress()));
        }
        CellRangeAddress res = new CellRangeAddress(address.getRow() - 1,
                address.getRow() - 1, address.getCol() - 2, address.getCol() - 1);
        List<MutableCellRangeAddress> intersectedRegion = findIntersectedRegion(res);

        addMergedRegion(res, intersectedRegion);
    }

    /**
     * Add merged region.
     * Static merge or iteration merge.
     * @param res merge region
     */
    public void addMergedRegion(CellRangeAddress res) {
        List<MutableCellRangeAddress> intersectedRegion = findIntersectedRegion(res);
        addMergedRegion(res, intersectedRegion);
    }

    private void addMergedRegion(CellRangeAddress res, List<MutableCellRangeAddress> intersectedRegion) {
        if (intersectedRegion.size() == 1) {
            mergeIntersectedRegions(intersectedRegion.get(0), res);
        } else if (intersectedRegion.isEmpty()) {
            mergedRegions.add(new MutableCellRangeAddress(res));
        } else {
            mergeIntersectedRegions(intersectedRegion, res);
        }
    }

    Stream<CellRangeAddress> getMergedRegions(){
        return mergedRegions.stream().map(MutableCellRangeAddress::toCellRangeAddress);
    }

    /**
     * Find main region with current cell.
     * For defining value in region.
     * @param res range of current merge region.
     * @return Current merge region if intersects is false; otherwise - intersected existing region.
     */
    public CellRangeAddress findIntersectedRange(CellRangeAddress res) {
        if (res.getFirstColumn() < 0 || res.getFirstRow() < 0) {
            throw new IllegalArgumentException(
                    String.format("Cannot merge first row with upper cell or first column with left cell: %s",
                            res.formatAsString()));
        }
        List<MutableCellRangeAddress> intersected = findIntersectedRegion(res);
        return intersected.stream().map(MutableCellRangeAddress::toCellRangeAddress).findFirst().orElse(res);
    }

    /**
     * Clear list.
     */
    public void clear() {
        mergedRegions.clear();
    }

    private void mergeIntersectedRegions(List<MutableCellRangeAddress> regions, CellRangeAddress res) {
        Comparator<MutableCellRangeAddress> minComp = (a, b) -> {
            int yDiff = a.getFirstRow() - b.getFirstRow();
            if (yDiff == 0) {
                return a.getFirstColumn() - b.getFirstColumn();
            }
            return yDiff;
        };
        Comparator<MutableCellRangeAddress> maxComp = (a, b) -> {
            int yDiff = a.getLastRow() - b.getLastRow();
            if (yDiff == 0) {
                return a.getLastColumn() - b.getLastColumn();
            }
            return yDiff;
        };

        MutableCellRangeAddress leftTop = regions.stream().min(minComp).get();
        MutableCellRangeAddress rightBottom = regions.stream().max(maxComp).get();

        if (res.getLastRow() == rightBottom.getLastRow() && res.getLastColumn() == rightBottom.getLastColumn()
                && res.getLastRow() > leftTop.getFirstRow() && res.getLastColumn() > leftTop.getFirstColumn()) {

            MutableCellRangeAddress unionedRegion = new MutableCellRangeAddress(
                    new CellRangeAddress(leftTop.getFirstRow(), rightBottom.getLastRow(),
                            leftTop.getFirstColumn(), rightBottom.getLastColumn()));

            this.mergedRegions.removeAll(regions);
            this.mergedRegions.add(unionedRegion);
        }
    }

    /**
     * Merge current cell with existing region.
     * @param mergedRegion existing merged region.
     * @param res current cell with it's neighbour (up or left)
     */
    private void mergeIntersectedRegions(MutableCellRangeAddress mergedRegion, CellRangeAddress res) {
        if (isValidUnion(mergedRegion.toCellRangeAddress(), res)) {
            CellRangeAddress unionRangeRes = new CellRangeAddress(mergedRegion.getFirstRow(), res.getLastRow(),
                    mergedRegion.getFirstColumn(), res.getLastColumn());
            mergedRegion.setAddress(unionRangeRes);
        }
    }

    /**
     * Check if is it valid union.
     * @param mergedRegion existing merged region
     * @param result region that will be concatenated.
     * @return True if it is one row (or column); otherwise - false.
     */
    private boolean isValidUnion(CellRangeAddress mergedRegion, CellRangeAddress result) {
        int startRow = mergedRegion.getFirstRow();
        int startCol = mergedRegion.getFirstColumn();
        int finishRow = mergedRegion.getLastRow();
        int finishCol = mergedRegion.getLastColumn();

        return (result.getFirstRow() == startRow && result.getLastRow() == finishRow)
                || (result.getFirstColumn() == startCol && result.getLastColumn() == finishCol);
    }

    /**
     * Find intersected existing region.
     * @param res range address
     * @return optional
     */
    private List<MutableCellRangeAddress> findIntersectedRegion(CellRangeAddress res) {
        return mergedRegions
                .stream()
                .filter(addr -> addr.intersects(res))
                .collect(Collectors.toList());
    }

    /**
     * Inner wrapper-class for CellRangeAddress.
     */
    private static class MutableCellRangeAddress {
        private CellRangeAddress address;

        /**
         * Simple constructor.
         * @param address address
         */
        MutableCellRangeAddress(CellRangeAddress address) {
            this.address = address;
        }

        /**
         * Check if regions is intersected.
         * @param addressNewRegion new region.
         * @return true if intersected; otherwise - false;
         */
        private boolean intersects(CellRangeAddress addressNewRegion) {
            return this.address.intersects(addressNewRegion);
        }

        /**
         * Setter for address.
         * @param address address.
         */
        private void setAddress(CellRangeAddress address) {
            this.address = address;
        }

        /**
         * Get cellRangeAddress.
         * @return cell range address
         */
        private CellRangeAddress toCellRangeAddress() {
            return this.address;
        }

        /**
         * Get first row.
         * @return first row
         */
        private int getFirstRow() {
            return address.getFirstRow();
        }

        /**
         * Get first column.
         * @return get first column
         */
        private int getFirstColumn() {
            return address.getFirstColumn();
        }

        /**
         * Get last row.
         * @return last row
         */
        private int getLastRow() {
            return address.getLastRow();
        }

        /**
         * Get last column.
         * @return last column
         */
        private int getLastColumn() {
            return address.getLastColumn();
        }
    }
}
