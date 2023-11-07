package com.github.chengyuxing.excel.type;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel header builder.
 */
public class XHeader {
    private final List<XRow> rows = new ArrayList<>();
    private int maxRowNumber = 0;
    private int maxColumnNumber = 0;

    /**
     * Add one row.
     *
     * @param row row
     * @return XHeader
     */
    public XHeader add(XRow row) {
        if (!isEmpty()) {
            XRow lastRow = rows.get(rows.size() - 1);
            int maxLastRow = 0;
            for (String field : lastRow.getFields()) {
                CellRangeAddress cellAddresses = lastRow.getCellAddresses(field);
                if (cellAddresses.getLastRow() > maxLastRow) {
                    maxLastRow = cellAddresses.getLastRow();
                }
            }
            List<String> currentFields = row.getFields();
            for (String currentField : currentFields) {
                CellRangeAddress cellAddresses = row.getCellAddresses(currentField);
                if (cellAddresses.getFirstRow() == 0 && cellAddresses.getLastRow() == 0) {
                    int nextRowNumber = maxLastRow + 1;
                    cellAddresses.setFirstRow(nextRowNumber);
                    cellAddresses.setLastRow(nextRowNumber);
                }
            }
        }
        rows.add(row);
        return this;
    }

    public boolean isEmpty() {
        return rows.isEmpty();
    }

    /**
     * Get header max row number.
     *
     * @return max row number
     */
    public int getMaxRowNumber() {
        for (XRow xRow : rows) {
            int rowNumber = xRow.getMaxRowNumber();
            if (rowNumber > maxRowNumber) {
                maxRowNumber = rowNumber;
            }
        }
        return maxRowNumber;
    }

    /**
     * Get header max column number.
     *
     * @return max column number
     */
    public int getMaxColumnNumber() {
        for (XRow xRow : rows) {
            int columnNumber = xRow.getMaxColumnNumber();
            if (columnNumber > maxColumnNumber) {
                maxColumnNumber = columnNumber;
            }
        }
        return maxColumnNumber;
    }

    /**
     * Get header data rows.
     *
     * @return header data rows
     */
    public List<XRow> getRows() {
        return rows;
    }
}
