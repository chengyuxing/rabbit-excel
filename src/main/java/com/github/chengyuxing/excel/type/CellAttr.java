package com.github.chengyuxing.excel.type;

import com.github.chengyuxing.excel.style.StyleConfig;
import org.apache.poi.ss.util.CellRangeAddress;

public class CellAttr {
    private StyleConfig cellStyle;
    private CellRangeAddress cellRangeAddress;

    public StyleConfig getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(StyleConfig cellStyle) {
        this.cellStyle = cellStyle;
    }

    public CellRangeAddress getCellRangeAddress() {
        return cellRangeAddress;
    }

    public void setCellRangeAddress(CellRangeAddress cellRangeAddress) {
        this.cellRangeAddress = cellRangeAddress;
    }
}
