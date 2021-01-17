package rabbit.excel.io;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import rabbit.common.types.DataRow;
import rabbit.common.types.TiFunction;
import rabbit.excel.style.XStyle;
import rabbit.excel.type.XSheet;
import rabbit.excel.type.XHeader;
import rabbit.excel.type.XRow;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

/**
 * excel文件生成器
 */
public class ExcelWriter implements AutoCloseable {
    public final static Logger log = LoggerFactory.getLogger(ExcelWriter.class);
    private final Workbook workbook;
    private final List<XSheet> xSheets = new ArrayList<>();

    /**
     * Excel读取类构造函数
     *
     * @param workbook 工作薄
     */
    public ExcelWriter(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * 创建一个新的空白单元格样式
     *
     * @return 空白单元格样式
     * @see XStyle
     */
    public XStyle createStyle() {
        return new XStyle(workbook.createCellStyle());
    }

    /**
     * 创建一个新的空白字形
     *
     * @return 空白字形
     */
    public Font createFont() {
        return workbook.createFont();
    }

    /**
     * 写入sheet数据
     *
     * @param xSheet sheet数据
     * @param more   更多的sheet数据
     * @return Excel写入类
     */
    public ExcelWriter write(XSheet xSheet, XSheet... more) {
        xSheets.add(xSheet);
        xSheets.addAll(Arrays.asList(more));
        return this;
    }

    /**
     * 写入sheet数据
     *
     * @param xSheets 一组sheet数据
     * @return Excel写入类
     */
    public ExcelWriter write(Collection<XSheet> xSheets) {
        this.xSheets.addAll(xSheets);
        return this;
    }

    /**
     * 获取excel文件字节流
     *
     * @return 字节流
     */
    public byte[] toBytes() {
        if (xSheets.size() < 1) {
            throw new IllegalStateException("there is nothing to write! don't you invoke method write(...) to add sheet data?");
        }
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            for (XSheet s : xSheets) {
                Sheet sheet = workbook.createSheet(s.getName());
                writeSheet(sheet, s);
            }
            workbook.write(out);
        } catch (IOException e) {
            log.error("io ex:{}", e.getMessage());
        }
        return out.toByteArray();
    }

    /**
     * 写Excel到输出流
     *
     * @param outputStream 输出流
     * @param close        是否在完成后关闭输出流
     * @throws IOException ioEx
     */
    public void saveTo(OutputStream outputStream, boolean close) throws IOException {
        outputStream.write(toBytes());
        if (close) {
            outputStream.flush();
            outputStream.close();
        }
    }

    /**
     * 写Excel到输出流并关闭输出流
     *
     * @param outputStream 输出流
     * @throws IOException ioEx
     */
    public void saveTo(OutputStream outputStream) throws IOException {
        saveTo(outputStream, true);
    }

    /**
     * 保存Excel到指定路径下
     *
     * @param path 文件保存路径（后缀可选）
     * @throws IOException ioEx
     */
    public void saveTo(String path) throws IOException {
        String suffix = "";
        if (!path.endsWith(".xlsx") && !path.endsWith(".xls")) {
            suffix = ".xlsx";
            if (workbook instanceof HSSFWorkbook) {
                suffix = ".xls";
            }
        }
        saveTo(new FileOutputStream(path + suffix));
    }

    /**
     * 保存Excel到文件对象
     *
     * @param file 文件对象
     * @throws IOException ioEx
     */
    public void saveTo(File file) throws IOException {
        saveTo(new FileOutputStream(file));
    }

    /**
     * 保存Excel到路径对象
     *
     * @param path 路径对象
     * @throws IOException ioEx
     */
    public void saveTo(Path path) throws IOException {
        saveTo(Files.newOutputStream(path));
    }

    /**
     * 写入数据到一个Sheet中
     *
     * @param sheet  sheet
     * @param xSheet sheet数据
     */
    private void writeSheet(Sheet sheet, XSheet xSheet) {
        XHeader xHeader = xSheet.getXHeader();
        List<DataRow> data = xSheet.getData();
        if (data != null && !data.isEmpty()) {
            List<String> fields = buildHeaderSpecial(sheet, xHeader, data.get(0).getNames(), xSheet.getHeaderStyle());
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
                    setCellStyle(cell, data.get(i), fields.get(j), j, xSheet.getCellStyle());
                }
            }
        } else {
            buildHeaderSpecial(sheet, xHeader, Collections.emptyList(), xSheet.getHeaderStyle());
        }
        if (xHeader.isEmpty()) {
            if (data != null && !data.isEmpty()) {
                autoColumnWidth(sheet, data.get(0).size());
            }
        } else {
            autoColumnWidth(sheet, xHeader);
        }
    }

    /**
     * 设置单元格的值
     *
     * @param cell  单元格
     * @param value 值
     * @param other 候选值
     */
    private void setCellValue(Cell cell, Object value, String other) {
        if (value == null || value.equals("")) {
            cell.setCellValue(other);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    /**
     * 设置单元格样式
     *
     * @param cell      单元格
     * @param row       行
     * @param column    字段名
     * @param index     当前列号
     * @param styleFunc 样式回调函数
     */
    private void setCellStyle(Cell cell, DataRow row, String column, int index, TiFunction<DataRow, String, Integer, XStyle> styleFunc) {
        if (styleFunc != null) {
            XStyle style = styleFunc.apply(row, column, index);
            if (style != null) {
                style.init();
                cell.setCellStyle(style.getStyle());
            }
        }
    }

    /**
     * 自动设置复杂表头的宽度
     *
     * @param sheet   sheet
     * @param xHeader 表头
     */
    private void autoColumnWidth(Sheet sheet, XHeader xHeader) {
        for (XRow xRow : xHeader.getRows()) {
            for (String field : xRow.getFields()) {
                sheet.autoSizeColumn(xRow.getCellAddresses(field).getFirstColumn());
            }
        }
    }

    /**
     * 自动设置简单表头的宽度
     *
     * @param sheet       sheet
     * @param columnCount 总列数
     */
    private void autoColumnWidth(Sheet sheet, int columnCount) {
        for (int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * 构建默认的简单表头
     *
     * @param sheet               sheet
     * @param defaultHeaderFields 默认的表头字段
     * @param xStyle              样式
     * @return 字段集合
     */
    private List<String> buildHeaderDefault(Sheet sheet, List<String> defaultHeaderFields, XStyle xStyle) {
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < defaultHeaderFields.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(defaultHeaderFields.get(i));
            if (xStyle != null) {
                xStyle.init();
                cell.setCellStyle(xStyle.getStyle());
            }
        }
        return defaultHeaderFields;
    }

    /**
     * 构建复杂的表头
     *
     * @param sheet               sheet
     * @param xHeader             表头数据
     * @param defaultHeaderFields 默认的表头字段
     * @param xStyle              样式
     * @return 字段集合
     */
    private List<String> buildHeaderSpecial(Sheet sheet, XHeader xHeader, List<String> defaultHeaderFields, XStyle xStyle) {
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
                    xCellStyle.init();
                    cellStyle = xCellStyle.getStyle();
                } else if (xStyle != null) {
                    // row style
                    xStyle.init();
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
        xSheets.clear();
    }
}
