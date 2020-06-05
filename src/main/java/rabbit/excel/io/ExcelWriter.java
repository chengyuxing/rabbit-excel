package rabbit.excel.io;

import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import rabbit.common.types.DataRow;
import rabbit.common.utils.ReflectUtil;
import rabbit.excel.style.IStyle;
import rabbit.excel.type.Head;
import rabbit.excel.type.ISheet;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * excel写入类
 */
public class ExcelWriter implements AutoCloseable {
    public final static Logger log = LoggerFactory.getLogger(ExcelWriter.class);

    private final Workbook workbook;
    private final List<ISheet<?, ?>> iSheets = new ArrayList<>();

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
    public ExcelWriter write(ISheet<?, ?> iSheet, ISheet<?, ?>... more) {
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
    public ExcelWriter write(Collection<ISheet<?, ?>> iSheets) {
        this.iSheets.addAll(iSheets);
        return this;
    }

    /**
     * 获取excel文件字节流
     *
     * @return 字节流
     */
    public byte[] getBytes() {
        if (iSheets.size() < 1) {
            throw new IllegalStateException("there is noting to write! don't you invoke method write(...) to add sheet data?");
        }
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            for (ISheet<?, ?> s : iSheets) {
                Sheet sheet = workbook.createSheet(s.getName());
                ExcelWriter.writeSheet(sheet, s);
            }
            workbook.write(out);
            workbook.close();
        } catch (IOException e) {
            log.error("io ex:{}", e.getMessage());
        } catch (NoSuchMethodException e) {
            log.error("no such method:{}", e.getMessage());
        } catch (InvocationTargetException e) {
            log.error("this object invoke failed:{}", e.getMessage());
        } catch (NoSuchFieldException e) {
            log.error("no such field:{}", e.getMessage());
        } catch (IllegalAccessException e) {
            log.error("access failed:{}", e.getMessage());
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
        outputStream.write(getBytes());
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
        saveTo(new FileOutputStream(fixedPath(path)));
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
     * @throws NoSuchMethodException     NoSuchMethodException
     * @throws NoSuchFieldException      NoSuchFieldException
     * @throws IllegalAccessException    IllegalAccessException
     * @throws InvocationTargetException InvocationTargetException
     */
    @SuppressWarnings("unchecked")
    public static void writeSheet(Sheet sheet, ISheet<?, ?> iSheet) throws NoSuchMethodException, NoSuchFieldException, IllegalAccessException, InvocationTargetException {
        Map<String, String> mapper = iSheet.getMapper();
        if (iSheet.getData() != null && iSheet.getData().size() > 0) {
            if (Map.class.isAssignableFrom(iSheet.getClazz())) {
                writeSheetOfMap(sheet, (ISheet<?, String>) iSheet);
            } else if (List.class.isAssignableFrom(iSheet.getClazz())) {
                writeSheetOfList(sheet, (ISheet<?, Integer>) iSheet);
            } else if (DataRow.class.isAssignableFrom(iSheet.getClazz())) {
                writeSheetOfDataRow(sheet, (ISheet<?, String>) iSheet);
            } else {
                //最后一种可能默认为java bean
                writeSheetOfJavaBean(sheet, (ISheet<?, String>) iSheet);
            }
        } else {
            buildHeader(sheet, mapper, iSheet.getHeaderStyle());
        }
    }

    @SuppressWarnings("unchecked")
    private static <T> void writeSheetOfMap(Sheet sheet, ISheet<T, String> iSheet) {
        List<Map<Object, Object>> data = (List<Map<Object, Object>>) iSheet.getData();
        Map<String, String> mapper = iSheet.getMapper();
        if (mapper.isEmpty()) {
            mapper = data.get(0).keySet().stream().collect(Collectors.toMap(Object::toString, Object::toString));
        }
        Object[] fields = buildHeader(sheet, mapper, iSheet.getHeaderStyle());
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < fields.length; j++) {
                Cell cell = row.createCell(j);
                Object value = data.get(i).get(fields[j]);
                setCellValue(cell, value, iSheet.getEmptyColumn());
                setCellStyle(cell, i, (String) fields[j], iSheet);
            }
        }
        autoColumnWidth(sheet, fields);
    }

    @SuppressWarnings("unchecked")
    private static <T> void writeSheetOfDataRow(Sheet sheet, ISheet<T, String> iSheet) {
        List<DataRow> data = (List<DataRow>) iSheet.getData();
        Map<String, String> mapper = iSheet.getMapper();
        if (mapper.isEmpty()) {
            mapper = data.get(0).toMap(Object::toString);
        }
        Object[] fields = buildHeader(sheet, mapper, iSheet.getHeaderStyle());
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < fields.length; j++) {
                Cell cell = row.createCell(j);
                Object value = data.get(i).get(fields[j].toString());
                setCellValue(cell, value, iSheet.getEmptyColumn());
                setCellStyle(cell, i, (String) fields[j], iSheet);
            }
        }
        autoColumnWidth(sheet, fields);
    }

    @SuppressWarnings("unchecked")
    private static <T> void writeSheetOfList(Sheet sheet, ISheet<T, Integer> iSheet) {
        List<List<Object>> data = (List<List<Object>>) iSheet.getData();
        IStyle iStyle = iSheet.getHeaderStyle();
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i);
            if (i == 0) {
                for (int j = 0; j < data.get(0).size(); j++) {
                    Cell cell = row.createCell(j);
                    Object value = data.get(i).get(j);
                    setCellValue(cell, value, iSheet.getEmptyColumn());
                    if (iStyle != null) {
                        cell.setCellStyle(iStyle.getStyle());
                    }
                }
            } else {
                for (int j = 0; j < data.get(0).size(); j++) {
                    Cell cell = row.createCell(j);
                    Object value = data.get(i).get(j);
                    setCellValue(cell, value, iSheet.getEmptyColumn());
                    setCellStyle(cell, i, j, iSheet);
                }
            }
        }
        autoColumnWidth(sheet, data.get(0).toArray());
    }

    private static <T> void writeSheetOfJavaBean(Sheet sheet, ISheet<T, String> iSheet) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException, NoSuchFieldException {
        List<?> data = iSheet.getData();
        Map<String, String> mapper = iSheet.getMapper();
        if (mapper.isEmpty()) {
            mapper = getMapper(data.get(0).getClass());
        }
        Class<?> beanClass = data.get(0).getClass();
        Object[] fields = buildHeader(sheet, mapper, iSheet.getHeaderStyle());
        for (int i = 0; i < data.size(); i++) {
            Object instance = data.get(i);
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < fields.length; j++) {
                Cell cell = row.createCell(j);
                String field = fields[j].toString();
                String getMethod = ReflectUtil.initGetMethod(field, beanClass.getDeclaredField(field).getType());
                Object value = beanClass.getDeclaredMethod(getMethod).invoke(instance);
                setCellValue(cell, value, iSheet.getEmptyColumn());
                setCellStyle(cell, i, field, iSheet);
            }
        }
        autoColumnWidth(sheet, fields);
    }

    private static void setCellValue(Cell cell, Object value, String other) {
        if (value == null || value.equals("")) {
            cell.setCellValue(other);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    private static <T, U> void setCellStyle(Cell cell, int row, U column, ISheet<T, U> iSheet) {
        if (iSheet.getCellStyle() != null) {
            IStyle style = iSheet.getCellStyle().apply(iSheet.getData().get(row), column);
            if (style != null)
                cell.setCellStyle(style.getStyle());
        }
    }

    private static <T> Map<String, String> getMapper(Class<T> clazz) {
        Map<String, String> map = new LinkedHashMap<>();
        Stream.of(clazz.getDeclaredFields())
                .filter(f -> f.isAnnotationPresent(Head.class))
                .forEach(f -> {
                    Head head = f.getAnnotation(Head.class);
                    String value = head.value();
                    if (value.equals("")) {
                        map.put(f.getName(), f.getName());
                    } else {
                        map.put(f.getName(), value);
                    }
                });
        return map;
    }

    private static void autoColumnWidth(Sheet sheet, Object[] header) {
        for (int i = 0; i < header.length; i++) {
            sheet.autoSizeColumn(i, true);
        }
    }

    private static Object[] buildHeader(Sheet sheet, Map<String, String> mapper, IStyle iStyle) {
        Row headerRow = sheet.createRow(0);
        Object[] fields = mapper.keySet().toArray();
        for (int i = 0; i < fields.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(mapper.get(fields[i].toString()));
            if (iStyle != null)
                cell.setCellStyle(iStyle.getStyle());
        }
        return fields;
    }

    private static String fixedPath(String path) {
        if (!path.endsWith(".xlsx"))
            path += ".xlsx";
        return path;
    }

    @Override
    public void close() throws Exception {
        workbook.close();
        iSheets.clear();
    }
}
