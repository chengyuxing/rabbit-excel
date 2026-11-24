package com.github.chengyuxing.excel.type;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;
import java.util.stream.Collectors;

/**
 * Excel header builder.
 */
public class HeaderConfig {
    private final List<RowConfig> rows = new ArrayList<>();
    private int maxColumnNumber = 0;
    private int nextRowNumber = 0;

    /**
     * Add one row.
     *
     * @param row row
     * @return XHeader
     */
    public HeaderConfig add(RowConfig row) {
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

    void setNextRowNumber(int nextRowNumber) {
        this.nextRowNumber = nextRowNumber;
    }

    /**
     * Get header max column number.
     *
     * @return max column number
     */
    public int getMaxColumnNumber() {
        return maxColumnNumber;
    }

    void setMaxColumnNumber(int maxColumnNumber) {
        this.maxColumnNumber = maxColumnNumber;
    }

    /**
     * Get header data rows.
     *
     * @return header data rows
     */
    public List<RowConfig> getRows() {
        return rows;
    }

    /**
     * Generate HeaderConfig by data object without style.
     *
     * @param configData config data object
     * @return HeaderConfig
     */
    public static HeaderConfig fromData(HeaderConfigData configData) {
        HeaderConfig headerConfig = new HeaderConfig();
        headerConfig.setMaxColumnNumber(configData.getMaxColumnNumber());
        headerConfig.setNextRowNumber(configData.getNextRowNumber());
        configData.getRows().forEach(rd -> {
            RowConfig rowConfig = new RowConfig();
            rowConfig.setHasFieldMap(rd.isHasFieldMap());
            rowConfig.setMaxColumnNumber(rd.getMaxColumnNumber());
            rowConfig.setMaxRowNumber(rd.getMaxRowNumber());
            rd.getCells().forEach(ccd -> {
                        CellConfig cc = new CellConfig(ccd.getField(), ccd.getText(), CellRangeAddress.valueOf(ccd.getAddress()), null);
                        rowConfig.getCells().add(cc);
                    }
            );
            headerConfig.getRows().add(rowConfig);
        });
        return headerConfig;
    }

    /**
     * Convert the HeaderConfig to data for serializable.
     *
     * @return HeaderConfig data object
     */
    public HeaderConfigData toData() {
        HeaderConfigData headerConfigData = new HeaderConfigData();
        List<RowConfigData> data = rows.stream().map(rc -> {
            RowConfigData rowConfigData = new RowConfigData();
            List<CellConfigData> cells = rc.getCells().stream().map(cell -> {
                CellConfigData cellConfigData = new CellConfigData();
                cellConfigData.setAddress(cell.getAddress().formatAsString());
                cellConfigData.setField(cell.getField());
                cellConfigData.setText(cell.getText());
                return cellConfigData;
            }).collect(Collectors.toList());
            rowConfigData.setCells(cells);
            rowConfigData.setHasFieldMap(rc.isHasFieldMap());
            rowConfigData.setMaxColumnNumber(rc.getMaxColumnNumber());
            rowConfigData.setMaxRowNumber(rc.getMaxRowNumber());
            return rowConfigData;
        }).collect(Collectors.toList());
        headerConfigData.setRows(data);
        headerConfigData.setMaxColumnNumber(getMaxColumnNumber());
        headerConfigData.setNextRowNumber(getNextRowNumber());
        return headerConfigData;
    }

    public static class HeaderConfigData {
        private List<RowConfigData> rows = new ArrayList<>();
        private int maxColumnNumber = 0;
        private int nextRowNumber = 0;

        public List<RowConfigData> getRows() {
            return rows;
        }

        public void setRows(List<RowConfigData> rows) {
            this.rows = rows;
        }

        public int getMaxColumnNumber() {
            return maxColumnNumber;
        }

        public void setMaxColumnNumber(int maxColumnNumber) {
            this.maxColumnNumber = maxColumnNumber;
        }

        public int getNextRowNumber() {
            return nextRowNumber;
        }

        public void setNextRowNumber(int nextRowNumber) {
            this.nextRowNumber = nextRowNumber;
        }
    }

    public static class RowConfigData {
        private List<CellConfigData> cells = new ArrayList<>();
        private boolean hasFieldMap = false;
        private int maxRowNumber = 0;
        private int maxColumnNumber = 0;

        public List<CellConfigData> getCells() {
            return cells;
        }

        public void setCells(List<CellConfigData> cells) {
            this.cells = cells;
        }

        public boolean isHasFieldMap() {
            return hasFieldMap;
        }

        public void setHasFieldMap(boolean hasFieldMap) {
            this.hasFieldMap = hasFieldMap;
        }

        public int getMaxRowNumber() {
            return maxRowNumber;
        }

        public void setMaxRowNumber(int maxRowNumber) {
            this.maxRowNumber = maxRowNumber;
        }

        public int getMaxColumnNumber() {
            return maxColumnNumber;
        }

        public void setMaxColumnNumber(int maxColumnNumber) {
            this.maxColumnNumber = maxColumnNumber;
        }
    }

    public static class CellConfigData {
        private String field;
        private String text;
        private String address;

        public String getField() {
            return field;
        }

        public void setField(String field) {
            this.field = field;
        }

        public String getText() {
            return text;
        }

        public void setText(String text) {
            this.text = text;
        }

        public String getAddress() {
            return address;
        }

        public void setAddress(String address) {
            this.address = address;
        }
    }
}
