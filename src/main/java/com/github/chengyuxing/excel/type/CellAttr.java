package com.github.chengyuxing.excel.type;

import com.github.chengyuxing.excel.style.XStyle;
import org.apache.poi.ss.util.CellRangeAddress;

public class CellAttr {
    private XStyle cellStyle;
    private CellRangeAddress cellRangeAddress;

    public XStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(XStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public CellRangeAddress getCellRangeAddress() {
        return cellRangeAddress;
    }

    public void setCellRangeAddress(CellRangeAddress cellRangeAddress) {
        this.cellRangeAddress = cellRangeAddress;
    }
}
