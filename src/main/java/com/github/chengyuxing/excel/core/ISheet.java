package com.github.chengyuxing.excel.core;


import com.github.chengyuxing.common.types.DataRow;

import java.util.List;
import java.util.Map;

/**
 * Excel Sheet类
 */
public class ISheet {
    private String name;
    private Map<String, String> mapper;
    private List<?> data;
    private Class<?> clazz;
    private String fillEmpty;
//    private IStyle cellStyle;

    ISheet() {
    }

    /**
     * 创建一个行数据为标准javaBean类型的Sheet
     *
     * @param name sheet名
     * @param data 数据
     * @return sheet
     */
    public static ISheet ofJavaBean(String name, List<?> data) {
        ISheet sheet = new ISheet();
        sheet.setName(name);
        sheet.setData(data);
        sheet.setClazz(data.get(0).getClass());
        return sheet;
    }

    /**
     * 创建一个行数据为DataRow类型的Sheet
     *
     * @param name     sheet名
     * @param dataRows 数据
     * @param mapper   字段名映射
     * @return sheet
     */
    public static ISheet ofDataRow(String name, List<DataRow> dataRows, Map<String, String> mapper) {
        ISheet sheet = new ISheet();
        sheet.setName(name);
        sheet.setData(dataRows);
        sheet.setMapper(mapper);
        sheet.setClazz(DataRow.class);
        return sheet;
    }

    /**
     * 创建一个行数据为DataRow类型的Sheet
     *
     * @param name     sheet名
     * @param dataRows 数据
     * @return sheet
     */
    public static ISheet ofDataRow(String name, List<DataRow> dataRows) {
        return ofDataRow(name, dataRows, null);
    }

    /**
     * 创建一个行数据为序列类型的Sheet
     *
     * @param name sheet名
     * @param data 数据
     * @return sheet
     */
    public static ISheet ofList(String name, List<List<Object>> data) {
        return ofJavaBean(name, data);
    }

    /**
     * 创建一个行数据为Map类型的Sheet
     *
     * @param name   sheet名
     * @param data   数据
     * @param mapper 表头字段名映射[字段名,别名]
     * @return sheet
     */
    public static ISheet ofMap(String name, List<Map<String, Object>> data, Map<String, String> mapper) {
        ISheet sheet = new ISheet();
        sheet.setName(name);
        sheet.setData(data);
        sheet.setMapper(mapper);
        sheet.setClazz(data.get(0).getClass());
        return sheet;
    }

    /**
     * 创建一个行数据为Map类型的Sheet
     *
     * @param name sheet名
     * @param data 数据
     * @return sheet
     */
    public static ISheet ofMap(String name, List<Map<String, Object>> data) {
        return ofMap(name, data, null);
    }

    public List<?> getData() {
        return data;
    }

    void setData(List<?> data) {
        this.data = data;
    }

    void setClazz(Class<?> clazz) {
        this.clazz = clazz;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Map<String, String> getMapper() {
        return mapper;
    }

    public void setMapper(Map<String, String> mapper) {
        this.mapper = mapper;
    }

    public String getFillEmpty() {
        return fillEmpty;
    }

    public void setFillEmpty(String fillEmpty) {
        this.fillEmpty = fillEmpty;
    }
}