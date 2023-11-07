package com.github.chengyuxing.excel.type;

import com.github.chengyuxing.excel.style.XStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import com.github.chengyuxing.common.tuple.Triple;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel complex cell builder.
 */
public class XRow {
    private final List<String> fields = new ArrayList<>();
    private final List<Triple<String, CellRangeAddress, XStyle>> value = new ArrayList<>();
    private boolean hasFieldMap = false;
    private int i = 0;
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
    public XRow set(String field, String name, CellRangeAddress cellAddresses, XStyle cellStyle) {
        CellRangeAddress actuallyAddress;
        if (isEmpty()) {
            if (cellAddresses != null) {
                actuallyAddress = cellAddresses;
            } else {
                actuallyAddress = new CellRangeAddress(0, 0, 0, 0);
            }
        } else {
            if (cellAddresses != null) {
                actuallyAddress = cellAddresses;
            } else {
                CellRangeAddress lastAddress = value.get(value.size() - 1).getItem2();
                actuallyAddress = new CellRangeAddress(lastAddress.getFirstRow(), lastAddress.getFirstRow(), lastAddress.getLastColumn() + 1, lastAddress.getLastColumn() + 1);
            }
        }

        value.add(Triple.of(name, actuallyAddress, cellStyle));
        fields.add(field);
        if (!field.startsWith("#") && !field.endsWith("#")) {
            hasFieldMap = true;
        }
        return this;
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
        return set("#" + i++ + "#", name, cellAddresses, cellStyle);
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

    public boolean isEmpty() {
        return fields.isEmpty();
    }

    /**
     * Get data field index.
     *
     * @param field data field
     * @return field index
     */
    public int getIndex(String field) {
        return fields.indexOf(field);
    }

    /**
     * Get display name.
     *
     * @param field data field
     * @return display name
     */
    public String getName(String field) {
        int index = getIndex(field);
        return value.get(index).getItem1();
    }

    public CellRangeAddress getCellAddresses(String field) {
        int index = getIndex(field);
        return value.get(index).getItem2();
    }

    public XStyle getStyle(String field) {
        int index = getIndex(field);
        return value.get(index).getItem3();
    }

    /**
     * Get all fields.
     *
     * @return all fields
     */
    public List<String> getFields() {
        return fields;
    }

    public int getMaxRowNumber() {
        for (String field : fields) {
            CellRangeAddress cellAddresses = getCellAddresses(field);
            int rowNumber = cellAddresses.getLastRow();
            if (rowNumber > maxRowNumber) {
                maxRowNumber = rowNumber;
            }
        }
        return maxRowNumber;
    }

    public int getMaxColumnNumber() {
        for (String field : fields) {
            CellRangeAddress cellAddresses = getCellAddresses(field);
            int columnNumber = cellAddresses.getLastColumn();
            if (columnNumber > maxColumnNumber) {
                maxColumnNumber = columnNumber;
            }
        }
        return maxColumnNumber;
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
