package com.github.chengyuxing.excel.type;

import com.github.chengyuxing.excel.style.XStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jetbrains.annotations.NotNull;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel complex cell builder.
 */
public class XRow {
    private final List<XCell> cells = new ArrayList<>();
    private boolean hasFieldMap = false;
    private int maxRowNumber = 0;
    private int maxColumnNumber = 0;

    /**
     * Add data field map to header column display name.
     *
     * @param field         data field
     * @param name          display name
     * @param cellAddresses cell address like excel standard format<br>
     *                      e.g. {@code A1:F3}:
     *                      <blockquote>
     *                      <pre>CellRangeAddress.valueOf("A1:F3")</pre>
     *                      </blockquote>
     * @param cellStyle     cell style
     * @return current row
     */
    public XRow set(@NotNull String field, @NotNull String name, CellRangeAddress cellAddresses, XStyle cellStyle) {
        CellRangeAddress next = nextAddress(cellAddresses);

        maxRowNumber = Math.max(maxRowNumber, next.getLastRow());
        maxColumnNumber = Math.max(maxColumnNumber, next.getLastColumn());

        XCell cell = new XCell(field, name, next, cellStyle);
        cells.add(cell);

        if (isHasField(field)) {
            hasFieldMap = true;
        }
        return this;
    }

    protected CellRangeAddress nextAddress(CellRangeAddress specified) {
        if (specified != null) {
            return specified;
        }
        if (cells.isEmpty()) {
            return new CellRangeAddress(0, 0, 0, 0);
        }
        CellRangeAddress last = cells.get(cells.size() - 1).getAddress();
        return new CellRangeAddress(last.getFirstRow(),
                last.getFirstRow(),
                last.getLastColumn() + 1,
                last.getLastColumn() + 1);
    }

    /**
     * Add data field map to header column display name.
     *
     * @param field         data field
     * @param name          display name
     * @param cellAddresses cell address like excel standard format<br>
     *                      e.g. {@code A1:F3}:
     *                      <blockquote>
     *                      <pre>CellRangeAddress.valueOf("A1:F3")</pre>
     *                      </blockquote>
     * @return 当前行数据
     */
    public XRow set(String field, String name, CellRangeAddress cellAddresses) {
        return set(field, name, cellAddresses, null);
    }

    /**
     * Add data field map to header column display name.
     *
     * @param field     data field
     * @param name      display name
     * @param cellStyle cell style
     * @return current row
     */
    public XRow set(String field, String name, XStyle cellStyle) {
        return set(field, name, null, cellStyle);
    }

    /**
     * Add data field map to header column display name.
     *
     * @param field data field
     * @param name  display name
     * @return current row
     */
    public XRow set(String field, String name) {
        return set(field, name, null, null);
    }

    /**
     * Add header column display name.
     *
     * @param name          display name
     * @param cellAddresses cell address like excel standard format<br>
     *                      e.g. {@code A1:F3}:
     *                      <blockquote>
     *                      <pre>CellRangeAddress.valueOf("A1:F3")</pre>
     *                      </blockquote>
     * @param cellStyle     cell style
     * @return current row
     */
    public XRow add(String name, CellRangeAddress cellAddresses, XStyle cellStyle) {
        return set("#" + cells.size() + "#", name, cellAddresses, cellStyle);
    }

    /**
     * Add header column display name.
     *
     * @param name          display name
     * @param cellAddresses cell address like excel standard format<br>
     *                      e.g. {@code A1:F3}:
     *                      <blockquote>
     *                      <pre>CellRangeAddress.valueOf("A1:F3")</pre>
     *                      </blockquote>
     * @return current row
     */
    public XRow add(String name, CellRangeAddress cellAddresses) {
        return add(name, cellAddresses, null);
    }

    /**
     * Add header column display name.
     *
     * @param name      display name
     * @param cellStyle cell style
     * @return current row
     */
    public XRow add(String name, XStyle cellStyle) {
        return add(name, null, cellStyle);
    }

    /**
     * Add header column display name.
     *
     * @param name display name
     * @return current row
     */
    public XRow add(String name) {
        return add(name, null, null);
    }

    public void layoutAutoRows(int row) {
        for (XCell cell : cells) {
            if (cell.isAutoRow()) {
                CellRangeAddress a = cell.getAddress();
                a.setFirstRow(row);
                a.setLastRow(row);
            }
        }
    }

    public boolean isEmpty() {
        return cells.isEmpty();
    }

    public List<XCell> getCells() {
        return cells;
    }

    public int getMaxRowNumber() {
        return maxRowNumber;
    }

    public int getMaxColumnNumber() {
        return maxColumnNumber;
    }

    public boolean isHasField(String field) {
        return !field.startsWith("#") && !field.endsWith("#");
    }

    /**
     * Check data field has mapping with display name.
     *
     * @return true or false
     */
    public boolean isHasFieldMap() {
        return hasFieldMap;
    }
}
