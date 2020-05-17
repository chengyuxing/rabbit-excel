package com.github.chengyuxing.excel.utils;

import com.github.chengyuxing.common.types.DataRow;
import com.github.chengyuxing.common.utils.ReflectUtil;
import com.github.chengyuxing.excel.core.Head;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.InvocationTargetException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

/**
 * Excel工具类
 */
public final class ExcelUtils {
    public final static Logger log = LoggerFactory.getLogger(ExcelUtils.class);

    /**
     * 写入一组List到sheet
     *
     * @param sheet     sheet
     * @param data      数据
     * @param fillEmpty 填充空单元格
     */
    public static void writeSheetOfList(Sheet sheet, List<List<Object>> data, String fillEmpty) {
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < data.get(0).size(); j++) {
                Cell cell = row.createCell(j);
                if (i == 0) {
                    cell.setCellStyle(createStyle(sheet));
                }
                Object value = data.get(i).get(j);
                if (value == null || value.equals("")) {
                    cell.setCellValue(fillEmpty);
                } else {
                    cell.setCellValue(value.toString());
                }
            }
        }
        autoColumnWidth(sheet, data.get(0).toArray());
    }

    public static void writeSheetOfDataRow(Sheet sheet, List<DataRow> data, Map<String, String> mapper, String fillEmpty) {
        Object[] fields = buildHeader(sheet, mapper);
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < fields.length; j++) {
                Cell cell = row.createCell(j);
                Object value = data.get(i).get(fields[j].toString());
                if (value == null || value.equals("")) {
                    cell.setCellValue(fillEmpty);
                } else {
                    cell.setCellValue(value.toString());
                }
            }
        }
        autoColumnWidth(sheet, fields);
    }

    public static void writeSheetOfMap(Sheet sheet, List<Map<Object, Object>> data, Map<String, String> mapper, String fillEmpty) {
        Object[] fields = buildHeader(sheet, mapper);
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < fields.length; j++) {
                Cell cell = row.createCell(j);
                Object value = data.get(i).get(fields[j].toString());
                if (value == null || value.toString().trim().equals("")) {
                    cell.setCellValue(fillEmpty);
                } else {
                    cell.setCellValue(value.toString());
                }
            }
        }
        autoColumnWidth(sheet, fields);
    }

    public static void writeSheetOfJavaBean(Sheet sheet, List<?> data, String fillEmpty) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException, NoSuchFieldException {
        Class<?> beanClass = data.get(0).getClass();
        Map<String, String> mapper = getMapper(beanClass);
        Object[] fields = buildHeader(sheet, mapper);
        for (int i = 0; i < data.size(); i++) {
            Object instance = data.get(i);
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < fields.length; j++) {
                Cell cell = row.createCell(j);
                String field = fields[j].toString();
                String getMethod = ReflectUtil.initGetMethod(field, beanClass.getDeclaredField(field).getType());
                Object value = beanClass.getDeclaredMethod(getMethod).invoke(instance);
                if (value == null || value.toString().trim().equals("")) {
                    cell.setCellValue(fillEmpty);
                } else {
                    cell.setCellValue(value.toString());
                }
            }
        }
        autoColumnWidth(sheet, fields);
    }

    /**
     * 创建第一行并返回一组表头的字段
     *
     * @param sheet  sheet
     * @param mapper 映射
     * @return 表头
     */
    public static Object[] buildHeader(Sheet sheet, Map<String, String> mapper) {
        Row headerRow = sheet.createRow(0);
        Object[] fields = mapper.keySet().toArray();
        for (int i = 0; i < fields.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellStyle(createStyle(sheet));
            cell.setCellValue(mapper.get(fields[i].toString()));
        }
        return fields;
    }

    /**
     * 获取实体类字段和表头的映射
     *
     * @param clazz 　实体类
     * @param <T>   　类型参数
     * @return 映射
     */
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

    public static Object getValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return (long) cell.getNumericCellValue();
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

//    public static Object getValue(Cell cell) {
//        switch (cell.getCellType()) {
//            case CELL_TYPE_STRING:
//                return cell.getStringCellValue();
//            case CELL_TYPE_BOOLEAN:
//                return cell.getBooleanCellValue();
//            case CELL_TYPE_NUMERIC:
//                if (DateUtil.isCellDateFormatted(cell)) {
//                    return cell.getDateCellValue();
//                }
//                return (long) cell.getNumericCellValue();
//            case CELL_TYPE_FORMULA:
//                return cell.getCellFormula();
//            default:
//                return "";
//        }
//    }

    public static void autoColumnWidth(Sheet sheet, Object[] header) {
        for (int i = 0; i < header.length; i++) {
            sheet.autoSizeColumn(i, true);
        }
    }

    public static CellStyle createStyle(Sheet sheet) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFont(font);
        return style;
    }
}
