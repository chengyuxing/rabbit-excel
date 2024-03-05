package com.github.chengyuxing.excel.type;

import com.github.chengyuxing.common.DataRow;
import com.github.chengyuxing.common.TiFunction;
import com.github.chengyuxing.excel.style.XStyle;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel Sheet data container.
 */
public class XSheet {
    private String name;
    private XHeader xHeader;
    private List<DataRow> data;
    private String emptyColumn = "";
    private final Map<String, Integer> fieldColumnWidths = new HashMap<>();
    private final Map<Integer, Integer> indexColumnWidths = new HashMap<>();
    private TiFunction<DataRow, String, Coord, CellAttr> cellAttr;
    private XStyle headerStyle;

    XSheet() {
    }

    /**
     * Returns a sheet data container with initial args.
     *
     * @param name    sheet name
     * @param data    data
     * @param xHeader header
     * @return XSheet
     */
    public static XSheet of(String name, List<DataRow> data, XHeader xHeader) {
        XSheet sheet = new XSheet();
        sheet.setName(name);
        sheet.setData(data);
        sheet.setXHeader(xHeader);
        return sheet;
    }

    /**
     * Returns a sheet data container with initial args.
     *
     * @param name   sheet name
     * @param data   data
     * @param header header
     * @return XSheet
     */
    public static XSheet of(String name, List<DataRow> data, XRow header) {
        XSheet sheet = new XSheet();
        sheet.setName(name);
        sheet.setData(data);
        XHeader xHeader = new XHeader();
        xHeader.add(header);
        sheet.setXHeader(xHeader);
        return sheet;
    }

    /**
     * Returns a sheet data container with initial args.
     *
     * @param name sheet name
     * @param data data
     * @return XSheet
     */
    public static XSheet of(String name, List<DataRow> data) {
        return of(name, data, new XHeader());
    }

    public XSheet columnWidth(String field, int width) {
        this.fieldColumnWidths.put(field, width);
        return this;
    }

    public XSheet columnWidth(int index, int width) {
        this.indexColumnWidths.put(index, width);
        return this;
    }

    public void setFieldColumnWidths(Map<String, Integer> columnWidths) {
        this.fieldColumnWidths.putAll(columnWidths);
    }

    public void setIndexColumnWidths(Map<Integer, Integer> columnWidths) {
        this.indexColumnWidths.putAll(columnWidths);
    }

    public XStyle getHeaderStyle() {
        return headerStyle;
    }

    public void setHeaderStyle(XStyle headerStyle) {
        this.headerStyle = headerStyle;
    }

    public List<DataRow> getData() {
        return data;
    }

    void setData(List<DataRow> data) {
        this.data = data;
    }

    public String getName() {
        return name;
    }

    void setName(String name) {
        this.name = name;
    }

    public XHeader getXHeader() {
        return xHeader;
    }

    void setXHeader(XHeader xHeader) {
        this.xHeader = xHeader;
    }

    public String getEmptyColumn() {
        return emptyColumn;
    }

    public Map<String, Integer> getFieldColumnWidths() {
        return fieldColumnWidths;
    }

    public Map<Integer, Integer> getIndexColumnWidths() {
        return indexColumnWidths;
    }

    public void setEmptyColumn(String emptyColumn) {
        this.emptyColumn = emptyColumn;
    }

    public TiFunction<DataRow, String, Coord, CellAttr> getCellAttr() {
        return cellAttr;
    }

    public void setCellAttr(TiFunction<DataRow, String, Coord, CellAttr> cellAttr) {
        this.cellAttr = cellAttr;
    }
}