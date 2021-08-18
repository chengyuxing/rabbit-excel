package com.github.chengyuxing.excel.type;

import com.github.chengyuxing.common.DataRow;
import com.github.chengyuxing.common.TiFunction;
import com.github.chengyuxing.excel.style.XStyle;

import java.util.List;

/**
 * Excel Sheet数据类<br>
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
     * 创建一个sheet<br>
     *
     * @param name    名称
     * @param data    数据
     * @param xHeader 表头
     * @return sheet
     * @see DataRow
     */
    public static XSheet of(String name, List<DataRow> data, XHeader xHeader) {
        XSheet sheet = new XSheet();
        sheet.setName(name);
        sheet.setData(data);
        sheet.setXHeader(xHeader);
        return sheet;
    }

    /**
     * 创建一个sheet
     *
     * @param name   名称
     * @param data   数据
     * @param header 单行表头
     * @return sheet
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
     * 创建一个sheet<br>
     *
     * @param name 名称
     * @param data 数据
     * @return sheet
     * @see DataRow
     */
    public static XSheet of(String name, List<DataRow> data) {
        return of(name, data, new XHeader());
    }

    /**
     * 获取表头样式
     *
     * @return 表头样式
     */
    public XStyle getHeaderStyle() {
        return headerStyle;
    }

    /**
     * 设置表头样式
     *
     * @param headerStyle 表头样式
     */
    public void setHeaderStyle(XStyle headerStyle) {
        this.headerStyle = headerStyle;
    }

    /**
     * 获取表体单元格样式函数
     *
     * @return 表体单元格样式函数
     */
    public TiFunction<DataRow, String, Coord, XStyle> getCellStyle() {
        return cellStyle;
    }

    /**
     * 设置表体单元格样式函数<br><br>
     * e.g. c字段大于700则添加红框例子：
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
     * @param cellStyle 单元格样式回调函数 {@code <数据行，列名，单元格坐标>}
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