package rabbit.excel.io;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import rabbit.common.types.DataRow;
import rabbit.common.types.TiFunction;
import rabbit.excel.style.IStyle;
import rabbit.excel.type.ISheet;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;
import java.util.stream.Collectors;

/**
 * excel文件生成器
 */
public class ExcelWriter implements AutoCloseable {
    public final static Logger log = LoggerFactory.getLogger(ExcelWriter.class);
    private final Workbook workbook;
    private final List<ISheet> iSheets = new ArrayList<>();

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
     * @see IStyle
     */
    public CellStyle createCellStyle() {
        return workbook.createCellStyle();
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
     * @param iSheet sheet数据
     * @param more   更多的sheet数据
     * @return Excel写入类
     */
    public ExcelWriter write(ISheet iSheet, ISheet... more) {
        iSheets.add(iSheet);
        iSheets.addAll(Arrays.asList(more));
        return this;
    }

    /**
     * 写入sheet数据
     *
     * @param iSheets 一组sheet数据
     * @return Excel写入类
     */
    public ExcelWriter write(Collection<ISheet> iSheets) {
        this.iSheets.addAll(iSheets);
        return this;
    }

    /**
     * 获取excel文件字节流
     *
     * @return 字节流
     */
    public byte[] toBytes() {
        if (iSheets.size() < 1) {
            throw new IllegalStateException("there is nothing to write! don't you invoke method write(...) to add sheet data?");
        }
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            for (ISheet s : iSheets) {
                Sheet sheet = workbook.createSheet(s.getName());
                ExcelWriter.writeSheet(sheet, s);
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
     * @param iSheet sheet数据
     */
    private static void writeSheet(Sheet sheet, ISheet iSheet) {
        Map<String, String> mapper = iSheet.getMapper();
        List<DataRow> data = iSheet.getData();
        if (data != null && data.size() > 0) {
            if (mapper.isEmpty()) {
                mapper = data.get(0).getNames().stream().collect(Collectors.toMap(k -> k, v -> v));
            }
            String[] fields = buildHeader(sheet, mapper, iSheet.getHeaderStyle());
            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(i + 1);
                for (int j = 0; j < fields.length; j++) {
                    Cell cell = row.createCell(j);
                    Object value = data.get(i).get(fields[j]);
                    setCellValue(cell, value, iSheet.getEmptyColumn());
                    setCellStyle(cell, data.get(i), fields[j], j, iSheet.getCellStyle());
                }
            }
            autoColumnWidth(sheet, fields);
        } else {
            buildHeader(sheet, mapper, iSheet.getHeaderStyle());
        }
    }

    private static void setCellValue(Cell cell, Object value, String other) {
        if (value == null || value.equals("")) {
            cell.setCellValue(other);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    private static void setCellStyle(Cell cell, DataRow row, String column, int index, TiFunction<DataRow, String, Integer, IStyle> styleFunc) {
        if (styleFunc != null) {
            IStyle style = styleFunc.apply(row, column, index);
            if (style != null) {
                style.init();
                cell.setCellStyle(style.getStyle());
            }
        }
    }

    private static void autoColumnWidth(Sheet sheet, Object[] header) {
        for (int i = 0; i < header.length; i++) {
            sheet.autoSizeColumn(i, true);
        }
    }

    private static String[] buildHeader(Sheet sheet, Map<String, String> mapper, IStyle iStyle) {
        Row headerRow = sheet.createRow(0);
        String[] fields = mapper.keySet().toArray(new String[0]);
        for (int i = 0; i < fields.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(mapper.get(fields[i]));
            if (iStyle != null) {
                iStyle.init();
                cell.setCellStyle(iStyle.getStyle());
            }
        }
        return fields;
    }

    @Override
    public void close() throws Exception {
        workbook.close();
        iSheets.clear();
    }
}
