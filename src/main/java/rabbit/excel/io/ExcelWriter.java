package rabbit.excel.io;

import org.apache.poi.ss.usermodel.*;
import rabbit.common.types.DataRow;
import rabbit.common.utils.ReflectUtil;
import rabbit.excel.types.Head;
import rabbit.excel.types.ISheet;

import java.lang.reflect.InvocationTargetException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * excel写入类
 */
public class ExcelWriter {
    private Workbook workbook;

    @SuppressWarnings("unchecked")
    public static void writeSheet(Sheet sheet, ISheet iSheet) throws NoSuchMethodException, NoSuchFieldException, IllegalAccessException, InvocationTargetException {
        Map<String, String> mapper = iSheet.getMapper();
        if (iSheet.getData() != null && iSheet.getData().size() > 0) {
            if (Map.class.isAssignableFrom(iSheet.getClazz())) {
                List<Map<Object, Object>> data = (List<Map<Object, Object>>) iSheet.getData();
                writeSheetOfMap(sheet, data, mapper, iSheet.getEmptyColumn());
            } else if (List.class.isAssignableFrom(iSheet.getClazz())) {
                List<List<Object>> data = (List<List<Object>>) iSheet.getData();
                writeSheetOfList(sheet, data, mapper, iSheet.getEmptyColumn());
            } else if (DataRow.class.isAssignableFrom(iSheet.getClazz())) {
                List<DataRow> data = (List<DataRow>) iSheet.getData();
                writeSheetOfDataRow(sheet, data, mapper, iSheet.getEmptyColumn());
            } else {
                //最后一种可能默认为java bean
                List<?> data = iSheet.getData();
                writeSheetOfJavaBean(sheet, data, mapper, iSheet.getEmptyColumn());
            }
        } else {
            buildHeader(sheet, mapper);
        }
    }

    private static void writeSheetOfMap(Sheet sheet, List<Map<Object, Object>> data, Map<String, String> mapper, String fillEmpty) {
        if (mapper.isEmpty()) {
            mapper = data.get(0).keySet().stream().collect(Collectors.toMap(Object::toString, Object::toString));
        }
        Object[] fields = buildHeader(sheet, mapper);
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < fields.length; j++) {
                Cell cell = row.createCell(j);
                Object value = data.get(i).get(fields[j]);
                if (value == null || value.toString().trim().equals("")) {
                    cell.setCellValue(fillEmpty);
                } else {
                    cell.setCellValue(value.toString());
                }
            }
        }
        autoColumnWidth(sheet, fields);
    }

    private static void writeSheetOfDataRow(Sheet sheet, List<DataRow> data, Map<String, String> mapper, String fillEmpty) {
        if (mapper.isEmpty()) {
            mapper = data.get(0).toMap(Object::toString);
        }
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

    private static void writeSheetOfList(Sheet sheet, List<List<Object>> data, Map<String, String> mapper, String fillEmpty) {
        int start = 0;
        if (!mapper.isEmpty()) {
            buildHeader(sheet, mapper);
            start = 1;
        }
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + start);
            for (int j = 0; j < data.get(0).size(); j++) {
                Cell cell = row.createCell(j);
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

    private static void writeSheetOfJavaBean(Sheet sheet, List<?> data, Map<String, String> mapper, String fillEmpty) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException, NoSuchFieldException {
        if (mapper.isEmpty()) {
            mapper = getMapper(data.get(0).getClass());
        }
        Class<?> beanClass = data.get(0).getClass();
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

    private static Object[] buildHeader(Sheet sheet, Map<String, String> mapper) {
        Row headerRow = sheet.createRow(0);
        Object[] fields = mapper.keySet().toArray();
        for (int i = 0; i < fields.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(mapper.get(fields[i].toString()));
        }
        return fields;
    }
}
