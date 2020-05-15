package org.rabbit.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.rabbit.common.types.DataRow;
import org.rabbit.excel.core.ExcelReader;
import org.rabbit.excel.core.ISheet;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static org.rabbit.excel.utils.ExcelUtils.*;

/**
 * Excel文件读写操作类
 */
public final class Excels {
    /**
     * 读Excel
     *
     * @param stream 输入流
     * @return Excel对象
     */
    public static ExcelReader read(InputStream stream) {
        return new ExcelReader(stream);
    }

    /**
     * 读Excel
     *
     * @param name 文件名
     * @return Excel
     * @throws FileNotFoundException ex
     */
    public static ExcelReader read(String name) throws FileNotFoundException {
        return read(new FileInputStream(name));
    }

    /**
     * 读Excel
     *
     * @param fileBytes 文件字节
     * @return Excel
     */
    public static ExcelReader read(byte[] fileBytes) {
        return read(new ByteArrayInputStream(fileBytes));
    }

    /**
     * 写Excel
     *
     * @param iSheet Sheet
     * @param more   更多Sheet
     * @return 字节流
     */
    @SuppressWarnings("unchecked")
    public static byte[] write(ISheet iSheet, ISheet... more) {
        ISheet[] sheets = new ISheet[more.length + 1];
        sheets[0] = iSheet;
        System.arraycopy(more, 0, sheets, 1, more.length);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            for (ISheet s : sheets) {
                Sheet sheet = workbook.createSheet(s.getName());
                Class<?> elementType = s.getClazz();
                if (List.class.isAssignableFrom(elementType)) {
                    List<List<Object>> data = (List<List<Object>>) s.getData();
                    writeSheetOfList(sheet, data, s.getFillEmpty());
                } else if (Map.class.isAssignableFrom(elementType)) {
                    List<Map<Object, Object>> data = (List<Map<Object, Object>>) s.getData();
                    Map<String, String> mapper = s.getMapper();
                    if (mapper == null || mapper.isEmpty()) {
                        mapper = data.get(0).keySet().stream().collect(Collectors.toMap(Object::toString, Object::toString));
                    }
                    writeSheetOfMap(sheet, data, mapper, s.getFillEmpty());
                } else if (DataRow.class.isAssignableFrom(elementType)) {
                    List<DataRow> dataRows = (List<DataRow>) s.getData();
                    Map<String, String> mapper = s.getMapper();
                    if (mapper == null || mapper.isEmpty()) {
                        mapper = dataRows.get(0).getNames().stream().collect(Collectors.toMap(k -> k, v -> v));
                    }
                    writeSheetOfDataRow(sheet, dataRows, mapper, s.getFillEmpty());
                } else {
                    writeSheetOfJavaBean(sheet, s.getData(), s.getFillEmpty());
                }
            }
            workbook.write(out);
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
     * @param output 输出流
     * @param iSheet sheet
     * @param more   更多sheet
     * @throws IOException ex
     */
    public static void write(OutputStream output, ISheet iSheet, ISheet... more) throws IOException {
        output.write(write(iSheet, more));
        output.flush();
        output.close();
    }

    /**
     * 写Excel到文件
     *
     * @param path   输出路径(包括文件名)
     * @param iSheet sheet
     * @param more   更多sheet
     * @throws IOException ex
     */
    public static void write(String path, ISheet iSheet, ISheet... more) throws IOException {
        write(new FileOutputStream(fixedPath(path)), iSheet, more);
    }

    private static String fixedPath(String path) {
        if (!path.endsWith(".xlsx"))
            path += ".xlsx";
        return path;
    }
}
