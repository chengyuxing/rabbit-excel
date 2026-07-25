package com.github.chengyuxing.excel.type;

import com.github.chengyuxing.excel.style.XStyle;
import org.apache.poi.ss.util.CellRangeAddress;

public class XCell {
    private final String field;
    private final String text;
    private final CellRangeAddress address;
    private final XStyle style;
    private boolean autoRow = false;

    public XCell(String field, String text, CellRangeAddress address, XStyle style) {
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

    public XStyle getStyle() {
        return style;
    }

    public boolean isAutoRow() {
        return autoRow;
    }
}
