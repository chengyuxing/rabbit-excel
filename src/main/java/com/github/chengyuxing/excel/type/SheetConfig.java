package com.github.chengyuxing.excel.type;

import com.github.chengyuxing.common.DataRow;
import com.github.chengyuxing.common.TiFunction;
import com.github.chengyuxing.excel.style.StyleConfig;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel Sheet data container.
 */
public class SheetConfig {
    private String name;
    private HeaderConfig headerConfig;
    private List<DataRow> data;
    private final Map<String, Integer> fieldColumnWidths = new HashMap<>();
    private final Map<Integer, Integer> indexColumnWidths = new HashMap<>();
    private TiFunction<DataRow, String, Coord, CellAttr> cellAttr;
    private StyleConfig headerStyle;

    SheetConfig() {
    }

    /**
     * Returns a sheet data container with initial args.
     *
     * @param name   sheet name
     * @param data   data
     * @param header header
     * @return XSheet
     */
    public static SheetConfig of(String name, List<DataRow> data, HeaderConfig header) {
        SheetConfig sheet = new SheetConfig();
        sheet.setName(name);
        sheet.setData(data);
        sheet.setXHeader(header);
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
    public static SheetConfig of(String name, List<DataRow> data, RowConfig header) {
        HeaderConfig headerConfig = new HeaderConfig();
        headerConfig.add(header);
        return of(name, data, headerConfig);
    }

    /**
     * Returns a sheet data container with initial args.
     *
     * @param name sheet name
     * @param data data
     * @return XSheet
     */
    public static SheetConfig of(String name, List<DataRow> data) {
        return of(name, data, new HeaderConfig());
    }

    public SheetConfig columnWidth(String field, int width) {
        this.fieldColumnWidths.put(field, width);
        return this;
    }

    public SheetConfig columnWidth(int index, int width) {
        this.indexColumnWidths.put(index, width);
        return this;
    }

    public void setFieldColumnWidths(Map<String, Integer> columnWidths) {
        this.fieldColumnWidths.putAll(columnWidths);
    }

    public void setIndexColumnWidths(Map<Integer, Integer> columnWidths) {
        this.indexColumnWidths.putAll(columnWidths);
    }

    public StyleConfig getHeaderStyle() {
        return headerStyle;
    }

    public void setHeaderStyle(StyleConfig headerStyle) {
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

    public HeaderConfig getXHeader() {
        return headerConfig;
    }

    void setXHeader(HeaderConfig headerConfig) {
        this.headerConfig = headerConfig;
    }

    public Map<String, Integer> getFieldColumnWidths() {
        return fieldColumnWidths;
    }

    public Map<Integer, Integer> getIndexColumnWidths() {
        return indexColumnWidths;
    }

    public TiFunction<DataRow, String, Coord, CellAttr> getCellAttr() {
        return cellAttr;
    }

    public void setCellAttr(TiFunction<DataRow, String, Coord, CellAttr> cellAttr) {
        this.cellAttr = cellAttr;
    }
}