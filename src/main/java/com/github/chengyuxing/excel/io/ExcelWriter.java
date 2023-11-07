package com.github.chengyuxing.excel.io;

import com.github.chengyuxing.common.DataRow;
import com.github.chengyuxing.common.TiFunction;
import com.github.chengyuxing.common.io.IOutput;
import com.github.chengyuxing.excel.style.XStyle;
import com.github.chengyuxing.excel.type.Coord;
import com.github.chengyuxing.excel.type.XHeader;
import com.github.chengyuxing.excel.type.XRow;
import com.github.chengyuxing.excel.type.XSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
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

    /**
     * {@inheritDoc}
     *
     * @return excel bytes
     */
    @Override
    public byte[] toBytes() throws IOException {
        if (xSheets.isEmpty()) {
            throw new IllegalStateException("there is nothing to write! don't you invoke method write(...) to add sheet data?");
        }
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        for (XSheet s : xSheets) {
            Sheet sheet = workbook.createSheet(s.getName());
            writeSheet(sheet, s);
        }
        workbook.write(out);
        return out.toByteArray();
    }

    /**
     * Save excel data to specify path.
     *
     * @param path file path (extension is optional)
     * @throws IOException ioEx
     */
    @Override
    public void saveTo(String path) throws IOException {
        String suffix = "";
        if (!path.endsWith(".xlsx") && !path.endsWith(".xls")) {
            suffix = ".xlsx";
            if (workbook instanceof HSSFWorkbook) {
                suffix = ".xls";
            }
        }
        saveTo(Files.newOutputStream(Paths.get(path + suffix)));
    }

    /**
     * Write data to sheet.
     *
     * @param sheet  sheet
     * @param xSheet sheet data container
     */
    void writeSheet(Sheet sheet, XSheet xSheet) {
        XHeader xHeader = xSheet.getXHeader();
        List<DataRow> data = xSheet.getData();
        if (data != null && !data.isEmpty()) {
            List<String> fields = buildHeaderSpecial(sheet, xHeader, data.get(0).names(), xSheet.getHeaderStyle());
            int columnCount = xHeader.getMaxColumnNumber() + 1;
            if (xHeader.isEmpty()) {
                columnCount = data.get(0).size();
            }
            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(xHeader.getMaxRowNumber() + 1 + i);
                for (int j = 0; j < columnCount; j++) {
                    Cell cell = row.createCell(j);
                    Object value = data.get(i).get(fields.get(j));
                    setCellValue(cell, value, xSheet.getEmptyColumn());
                    setCellStyle(cell, data.get(i), fields.get(j), new Coord(i, j), xSheet.getCellStyle());
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
    }

    void setCellValue(Cell cell, Object value, String other) {
        if (value == null || value.equals("")) {
            cell.setCellValue(other);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    void setCellStyle(Cell cell, DataRow row, String column, Coord coord, TiFunction<DataRow, String, Coord, XStyle> styleFunc) {
        if (styleFunc != null) {
            XStyle style = styleFunc.apply(row, column, coord);
            if (style != null) {
                cell.setCellStyle(style.getStyle());
            }
        }
    }

    void autoColumnWidth(Sheet sheet, XHeader xHeader) {
        for (XRow xRow : xHeader.getRows()) {
            for (String field : xRow.getFields()) {
                sheet.autoSizeColumn(xRow.getCellAddresses(field).getFirstColumn());
            }
        }
    }

    void autoColumnWidth(Sheet sheet, int columnCount) {
        for (int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    List<String> buildHeaderDefault(Sheet sheet, List<String> defaultHeaderFields, XStyle xStyle) {
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

    List<String> buildHeaderSpecial(Sheet sheet, XHeader xHeader, List<String> defaultHeaderFields, XStyle xStyle) {
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
            int startRow = xHeader.getMaxRowNumber() + 1;
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
        for (int i = 0; i <= xHeader.getMaxRowNumber(); i++) {
            sheet.createRow(i);
        }
        List<XRow> xRows = xHeader.getRows();
        for (XRow xRow : xRows) {
            List<String> keys = xRow.getFields();
            for (String key : keys) {
                CellRangeAddress cellAddresses = xRow.getCellAddresses(key);
                if (hasFieldMap && !key.startsWith("#") && !key.endsWith("#")) {
                    if (fields.length > cellAddresses.getFirstColumn()) {
                        fields[cellAddresses.getFirstColumn()] = key;
                    }
                }
                // merge columns first
                if (cellAddresses.getFirstColumn() != cellAddresses.getLastColumn() || cellAddresses.getFirstRow() != cellAddresses.getLastRow()) {
                    sheet.addMergedRegion(cellAddresses);
                }
                // get created row by actually row number
                Row headerRow = sheet.getRow(cellAddresses.getFirstRow());
                Cell cell = headerRow.createCell(cellAddresses.getFirstColumn());
                cell.setCellValue(xRow.getName(key));

                CellStyle cellStyle = null;
                // cell style first
                XStyle xCellStyle = xRow.getStyle(key);
                if (xCellStyle != null) {
                    cellStyle = xCellStyle.getStyle();
                } else if (xStyle != null) {
                    // row style
                    cellStyle = xStyle.getStyle();
                }
                cell.setCellStyle(cellStyle);
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
