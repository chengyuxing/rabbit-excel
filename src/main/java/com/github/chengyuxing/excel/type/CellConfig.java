package com.github.chengyuxing.excel.type;

import com.github.chengyuxing.excel.style.StyleConfig;
import org.apache.poi.ss.util.CellRangeAddress;

public class CellConfig {
    private final String field;
    private final String text;
    private final CellRangeAddress address;
    private final StyleConfig style;
    private final boolean autoRow;

    public CellConfig(String field, String text, CellRangeAddress address, StyleConfig style) {
        this.field = field;
        this.text = text;
        this.address = address;
        this.style = style;
        this.autoRow = address.getFirstRow() == 0 && address.getLastRow() == 0;
    }

    public CellRangeAddress getAddress() {
        return address;
    }

    public String getField() {
        return field;
    }

    public String getText() {
        return text;
    }

    public StyleConfig getStyle() {
        return style;
    }

    public boolean isAutoRow() {
        return autoRow;
    }
}
