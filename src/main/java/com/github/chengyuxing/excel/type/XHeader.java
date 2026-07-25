package com.github.chengyuxing.excel.type;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel header builder.
 */
public class XHeader {
    private final List<XRow> rows = new ArrayList<>();
    private int maxColumnNumber = 0;
    private int nextRowNumber = 0;

    /**
     * Add one row.
     *
     * @param row row
     * @return XHeader
     */
    public XHeader add(XRow row) {
        row.layoutAutoRows(nextRowNumber);
        nextRowNumber = Math.max(nextRowNumber, row.getMaxRowNumber() + 1);
        maxColumnNumber = Math.max(maxColumnNumber, row.getMaxColumnNumber());
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
    public int getNextRowNumber() {
        return nextRowNumber;
    }

    /**
     * Get header max column number.
     *
     * @return max column number
     */
    public int getMaxColumnNumber() {
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
