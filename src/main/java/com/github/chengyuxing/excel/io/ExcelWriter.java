package com.github.chengyuxing.excel.io;

import com.github.chengyuxing.common.DataRow;
import com.github.chengyuxing.common.TiFunction;
import com.github.chengyuxing.common.io.IOutput;
import com.github.chengyuxing.excel.style.StyleConfig;
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
    protected final List<SheetConfig> sheetConfigs = new ArrayList<>();

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
    public StyleConfig createStyle() {
        return new StyleConfig(workbook.createCellStyle());
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
     * @param sheetConfig sheet
     * @param more   more sheet
     * @return ExcelWriter
     */
    public ExcelWriter add(SheetConfig sheetConfig, SheetConfig... more) {
        sheetConfigs.add(sheetConfig);
        sheetConfigs.addAll(Arrays.asList(more));
        return this;
    }

    /**
     * Append more than one sheets ready to save.
     *
     * @param sheetConfigs sheets
     * @return ExcelWriter
     */
    public ExcelWriter add(Collection<SheetConfig> sheetConfigs) {
        this.sheetConfigs.addAll(sheetConfigs);
        return this;
    }

    @Override
    public void writeTo(OutputStream out) throws IOException {
        if (sheetConfigs.isEmpty()) {
            throw new IllegalStateException("there is nothing to write! don't you invoke method write(...) to add sheet data?");
        }
        for (SheetConfig s : sheetConfigs) {
            Sheet sheet = workbook.createSheet(s.getName());
            writeSheet(sheet, s);
        }
        workbook.setForceFormulaRecalculation(true);
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
     * @param sheetConfig sheet data container
     */
    protected void writeSheet(Sheet sheet, SheetConfig sheetConfig) {
        HeaderConfig headerConfig = sheetConfig.getXHeader();
        List<DataRow> data = sheetConfig.getData();
        if (data != null && !data.isEmpty()) {
            List<String> fields = buildHeaderSpecial(sheet, headerConfig, data.get(0).names(), sheetConfig.getHeaderStyle());
            int columnCount = headerConfig.getMaxColumnNumber() + 1;
            if (headerConfig.isEmpty()) {
                columnCount = data.get(0).size();
            }
            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(headerConfig.getNextRowNumber() + i);
                DataRow item = data.get(i);
                for (int j = 0; j < columnCount; j++) {
                    Cell cell = row.createCell(j);
                    String field = fields.get(j);
                    Object value = item.get(field);
                    setCellValue(cell, value);

                    TiFunction<DataRow, String, Coord, CellAttr> caFn = sheetConfig.getCellAttr();
                    if (caFn != null) {
                        CellAttr attr = caFn.apply(item, field, new Coord(i, j));
                        if (attr != null) {
                            CellRangeAddress address = attr.getCellRangeAddress();
                            if (address != null) {
                                sheet.addMergedRegion(address);
                            }
                            StyleConfig style = attr.getCellStyle();
                            if (style != null) {
                                cell.setCellStyle(style.getStyle());
                            }
                        }
                    }
                }
            }
        } else {
            buildHeaderSpecial(sheet, headerConfig, Collections.emptyList(), sheetConfig.getHeaderStyle());
        }
        // if big excel writer, do not set column width
        if (workbook instanceof SXSSFWorkbook) {
            return;
        }
        if (headerConfig.isEmpty()) {
            if (data != null && !data.isEmpty()) {
                autoColumnWidth(sheet, data.get(0).size());
            }
        } else {
            autoColumnWidth(sheet, headerConfig);
        }
        manualColumnWidth(sheet, sheetConfig);
    }

    protected void setCellValue(Cell cell, Object value) {
        ExcelUtils.setCellObject(cell, value);
    }

    protected void manualColumnWidth(Sheet sheet, SheetConfig sheetConfig) {
        HeaderConfig headerConfig = sheetConfig.getXHeader();
        Map<String, Integer> fieldWidths = sheetConfig.getFieldColumnWidths();
        Map<Integer, Integer> indexWidths = sheetConfig.getIndexColumnWidths();
        if (!fieldWidths.isEmpty()) {
            for (RowConfig rowConfig : headerConfig.getRows()) {
                for (CellConfig cellConfig : rowConfig.getCells()) {
                    if (fieldWidths.containsKey(cellConfig.getField())) {
                        int idx = cellConfig.getAddress().getFirstColumn();
                        int width = fieldWidths.get(cellConfig.getField());
                        sheet.setColumnWidth(idx, width);
                    }
                }
            }
        }
        if (!indexWidths.isEmpty()) {
            indexWidths.forEach(sheet::setColumnWidth);
        }
    }

    protected void autoColumnWidth(Sheet sheet, HeaderConfig headerConfig) {
        for (RowConfig rowConfig : headerConfig.getRows()) {
            for (CellConfig cellConfig : rowConfig.getCells()) {
                sheet.autoSizeColumn(cellConfig.getAddress().getFirstColumn());
            }
        }
    }

    protected void autoColumnWidth(Sheet sheet, int columnCount) {
        for (int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    protected List<String> buildHeaderDefault(Sheet sheet, List<String> defaultHeaderFields, StyleConfig styleConfig) {
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < defaultHeaderFields.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(defaultHeaderFields.get(i));
            if (styleConfig != null) {
                cell.setCellStyle(styleConfig.getStyle());
            }
        }
        return defaultHeaderFields;
    }

    protected List<String> buildHeaderSpecial(Sheet sheet, HeaderConfig headerConfig, List<String> defaultHeaderFields, StyleConfig styleConfig) {
        // just use DataRow's names default
        if (headerConfig.isEmpty()) {
            return buildHeaderDefault(sheet, defaultHeaderFields, styleConfig);
        }
        boolean hasFieldMap = false;
        for (RowConfig rowConfig : headerConfig.getRows()) {
            hasFieldMap = rowConfig.isHasFieldMap();
        }

        String[] fields = new String[0];
        // if has no field mapping relation, use DataRow's names as default
        if (!hasFieldMap) {
            RowConfig rowConfig = new RowConfig();
            int startRow = headerConfig.getNextRowNumber();
            if (!defaultHeaderFields.isEmpty()) {
                for (int i = 0; i < defaultHeaderFields.size(); i++) {
                    rowConfig.add(defaultHeaderFields.get(i), new CellRangeAddress(startRow, startRow, i, i));
                }
                headerConfig.add(rowConfig);
                fields = defaultHeaderFields.toArray(new String[0]);
            }
        } else {
            // maybe header's length > dataRow's length
            fields = new String[headerConfig.getMaxColumnNumber() + 1];
            Arrays.fill(fields, "___");
        }

        // total rows
        // create rows first.
        for (int i = 0; i <= headerConfig.getNextRowNumber(); i++) {
            sheet.createRow(i);
        }
        List<RowConfig> rowConfigs = headerConfig.getRows();
        for (RowConfig rowConfig : rowConfigs) {
            for (CellConfig cellConfig : rowConfig.getCells()) {
                String key = cellConfig.getField();
                CellRangeAddress address = cellConfig.getAddress();
                if (hasFieldMap && rowConfig.isHasField(cellConfig.getField())) {
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
                cell.setCellValue(cellConfig.getText());

                CellStyle style = null;
                // cell style first
                StyleConfig xCellStyle = cellConfig.getStyle();
                if (xCellStyle != null) {
                    style = xCellStyle.getStyle();
                } else if (styleConfig != null) {
                    // row style
                    style = styleConfig.getStyle();
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
        sheetConfigs.clear();
    }
}
