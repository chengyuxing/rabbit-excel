package com.github.chengyuxing.excel.type;

import com.github.chengyuxing.common.DataRow;
import com.github.chengyuxing.common.TiFunction;
import com.github.chengyuxing.excel.style.XStyle;

import java.util.List;

/**
 * Excel Sheet data container.
 */
public class XSheet {
    private String name;
    private XHeader xHeader;
    private List<DataRow> data;
    private String emptyColumn = "";
    private TiFunction<DataRow, String, Coord, XStyle> cellStyle;
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

    public XStyle getHeaderStyle() {
        return headerStyle;
    }

    public void setHeaderStyle(XStyle headerStyle) {
        this.headerStyle = headerStyle;
    }

    public TiFunction<DataRow, String, Coord, XStyle> getCellStyle() {
        return cellStyle;
    }

    /**
     * Set cell style function.<br>
     * e.g. set c field mapped cell red border if c field value {@code > } 700:
     * <blockquote>
     * <pre>
     *     XStyle danger = writer.createStyle();
     *     danger.setBorder(new Border(BorderStyle.DOUBLE, IndexedColors.RED));
     *     XSheet xSheet = ISheet.of("sheet1", list);
     *     xSheet.setCellStyle((row, key, coord) {@code ->} {
     *         if (key.equals("c"){@code &&} (double) row.get("c") {@code >} 700) {
     *             return danger;
     *         }
     *         return null;
     *     });</pre>
     * </blockquote>
     *
     * @param cellStyle (row, field, coord) {@code ->} style
     */
    public void setCellStyle(TiFunction<DataRow, String, Coord, XStyle> cellStyle) {
        this.cellStyle = cellStyle;
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

    public void setEmptyColumn(String emptyColumn) {
        this.emptyColumn = emptyColumn;
    }
}