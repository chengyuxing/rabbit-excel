package com.github.chengyuxing.excel.io;

import com.github.chengyuxing.common.DataRow;
import com.github.chengyuxing.common.TiFunction;
import com.github.chengyuxing.common.io.IOutput;
import com.github.chengyuxing.excel.style.XStyle;
import com.github.chengyuxing.excel.type.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

/**
 * Excel file writer.
 */
public class ExcelWriter implements IOutput, AutoCloseable {
    protected final Workbook workbook;
    protected final List<XSheet> xSheets = new ArrayList<>();

    /**
     * Constructs an ExcelWriter with Workbook.
     *
     * @param workbook workbook
     */
    public ExcelWriter(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * Create an empty cell style type.
     *
     * @return empty cell style
     */
    public XStyle createStyle() {
        return new XStyle(workbook.createCellStyle());
    }

    /**
     * Create an empty font.
     *
     * @return empty font
     */
    public Font createFont() {
        return workbook.createFont();
    }

    /**
     * Append more than one sheets ready to save.
     *
     * @param xSheet sheet
     * @param more   more sheet
     * @return ExcelWriter
     */
    public ExcelWriter write(XSheet xSheet, XSheet... more) {
        xSheets.add(xSheet);
        xSheets.addAll(Arrays.asList(more));
        return this;
    }

    /**
     * Append more than one sheets ready to save.
     *
     * @param xSheets sheets
     * @return ExcelWriter
     */
    public ExcelWriter write(Collection<XSheet> xSheets) {
        this.xSheets.addAll(xSheets);
        return this;
    }

    @Override
    public void writeTo(OutputStream out) throws IOException {
        if (xSheets.isEmpty()) {
            throw new IllegalStateException("there is nothing to write! don't you invoke method write(...) to add sheet data?");
        }
        for (XSheet s : xSheets) {
            Sheet sheet = workbook.createSheet(s.getName());
            writeSheet(sheet, s);
        }
        workbook.write(out);
    }

    /**
     * Save excel data to specify path.
     *
     * @param path file path (extension is optional)
     * @throws IOException ioEx
     */
    @Override
    public void writeTo(String path) throws IOException {
        String suffix = "";
        if (!path.endsWith(".xlsx") && !path.endsWith(".xls")) {
            suffix = ".xlsx";
            if (workbook instanceof HSSFWorkbook) {
                suffix = ".xls";
            }
        }
        IOutput.super.writeTo(path + suffix);
    }

    /**
     * Write data to sheet.
     *
     * @param sheet  sheet
     * @param xSheet sheet data container
     */
    protected void writeSheet(Sheet sheet, XSheet xSheet) {
        XHeader xHeader = xSheet.getXHeader();
        List<DataRow> data = xSheet.getData();
        if (data != null && !data.isEmpty()) {
            List<String> fields = buildHeaderSpecial(sheet, xHeader, data.get(0).names(), xSheet.getHeaderStyle());
            int columnCount = xHeader.getMaxColumnNumber() + 1;
            if (xHeader.isEmpty()) {
                columnCount = data.get(0).size();
            }
            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(xHeader.getNextRowNumber() + i);
                DataRow item = data.get(i);
                for (int j = 0; j < columnCount; j++) {
                    Cell cell = row.createCell(j);
                    String field = fields.get(j);
                    Object value = item.get(field);
                    setCellValue(cell, value, xSheet.getEmptyColumn());

                    TiFunction<DataRow, String, Coord, CellAttr> caFn = xSheet.getCellAttr();
                    if (caFn != null) {
                        CellAttr attr = caFn.apply(item, field, new Coord(i, j));
                        if (attr != null) {
                            CellRangeAddress address = attr.getCellRangeAddress();
                            if (address != null) {
                                sheet.addMergedRegion(address);
                            }
                            XStyle style = attr.getCellStyle();
                            if (style != null) {
                                cell.setCellStyle(style.getStyle());
                            }
                        }
                    }
                }
            }
        } else {
            buildHeaderSpecial(sheet, xHeader, Collections.emptyList(), xSheet.getHeaderStyle());
        }
        // if big excel writer, do not set column width
        if (workbook instanceof SXSSFWorkbook) {
            return;
        }
        if (xHeader.isEmpty()) {
            if (data != null && !data.isEmpty()) {
                autoColumnWidth(sheet, data.get(0).size());
            }
        } else {
            autoColumnWidth(sheet, xHeader);
        }
        manualColumnWidth(sheet, xSheet);
    }

    protected void setCellValue(Cell cell, Object value, String other) {
        if (value == null || value.equals("")) {
            cell.setCellValue(other);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    protected void manualColumnWidth(Sheet sheet, XSheet xSheet) {
        XHeader xHeader = xSheet.getXHeader();
        Map<String, Integer> fieldWidths = xSheet.getFieldColumnWidths();
        Map<Integer, Integer> indexWidths = xSheet.getIndexColumnWidths();
        if (!fieldWidths.isEmpty()) {
            for (XRow xRow : xHeader.getRows()) {
                for (XCell xCell : xRow.getCells()) {
                    if (fieldWidths.containsKey(xCell.getField())) {
                        int idx = xCell.getAddress().getFirstColumn();
                        int width = fieldWidths.get(xCell.getField());
                        sheet.setColumnWidth(idx, width);
                    }
                }
            }
        }
        if (!indexWidths.isEmpty()) {
            indexWidths.forEach(sheet::setColumnWidth);
        }
    }

    protected void autoColumnWidth(Sheet sheet, XHeader xHeader) {
        for (XRow xRow : xHeader.getRows()) {
            for (XCell xCell : xRow.getCells()) {
                sheet.autoSizeColumn(xCell.getAddress().getFirstColumn());
            }
        }
    }

    protected void autoColumnWidth(Sheet sheet, int columnCount) {
        for (int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    protected List<String> buildHeaderDefault(Sheet sheet, List<String> defaultHeaderFields, XStyle xStyle) {
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < defaultHeaderFields.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(defaultHeaderFields.get(i));
            if (xStyle != null) {
                cell.setCellStyle(xStyle.getStyle());
            }
        }
        return defaultHeaderFields;
    }

    protected List<String> buildHeaderSpecial(Sheet sheet, XHeader xHeader, List<String> defaultHeaderFields, XStyle xStyle) {
        // just use DataRow's names default
        if (xHeader.isEmpty()) {
            return buildHeaderDefault(sheet, defaultHeaderFields, xStyle);
        }
        boolean hasFieldMap = false;
        for (XRow xRow : xHeader.getRows()) {
            hasFieldMap = xRow.isHasFieldMap();
        }

        String[] fields = new String[0];
        // if has no field mapping relation, use DataRow's names as default
        if (!hasFieldMap) {
            XRow xRow = new XRow();
            int startRow = xHeader.getNextRowNumber();
            if (!defaultHeaderFields.isEmpty()) {
                for (int i = 0; i < defaultHeaderFields.size(); i++) {
                    xRow.add(defaultHeaderFields.get(i), new CellRangeAddress(startRow, startRow, i, i));
                }
                xHeader.add(xRow);
                fields = defaultHeaderFields.toArray(new String[0]);
            }
        } else {
            // maybe header's length > dataRow's length
            fields = new String[xHeader.getMaxColumnNumber() + 1];
            Arrays.fill(fields, "___");
        }

        // total rows
        // create rows first.
        for (int i = 0; i <= xHeader.getNextRowNumber(); i++) {
            sheet.createRow(i);
        }
        List<XRow> xRows = xHeader.getRows();
        for (XRow xRow : xRows) {
            for (XCell xCell : xRow.getCells()) {
                String key = xCell.getField();
                CellRangeAddress address = xCell.getAddress();
                if (hasFieldMap && xRow.isHasField(xCell.getField())) {
                    if (fields.length > address.getFirstColumn()) {
                        fields[address.getFirstColumn()] = key;
                    }
                }
                // merge columns first
                if (address.getFirstColumn() != address.getLastColumn() || address.getFirstRow() != address.getLastRow()) {
                    sheet.addMergedRegion(address);
                }
                // get created row by actually row number
                Row headerRow = sheet.getRow(address.getFirstRow());
                Cell cell = headerRow.createCell(address.getFirstColumn());
                cell.setCellValue(xCell.getText());

                CellStyle style = null;
                // cell style first
                XStyle xCellStyle = xCell.getStyle();
                if (xCellStyle != null) {
                    style = xCellStyle.getStyle();
                } else if (xStyle != null) {
                    // row style
                    style = xStyle.getStyle();
                }
                cell.setCellStyle(style);
            }
        }
        return Arrays.asList(fields);
    }

    @Override
    public void close() throws Exception {
        workbook.close();
        if (workbook instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook) workbook).dispose();
        }
        xSheets.clear();
    }
}
