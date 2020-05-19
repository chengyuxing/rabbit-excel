package rabbit.excel.types;


import rabbit.common.types.DataRow;

import java.util.Collections;
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
    private String emptyColumn = "";

    ISheet() {
    }

    /**
     * 创建一个sheet
     *
     * @param name   名称
     * @param data   数据类型参数支持：List&lt;Object&gt;; DataRow; Map&lt;String,Object&gt;; 标准Java Bean
     * @param mapper 表头字段名称映射(字段名，列名)
     * @return sheet
     * @see DataRow
     */
    public static ISheet of(String name, List<?> data, Map<String, String> mapper) {
        ISheet sheet = new ISheet();
        sheet.setName(name);
        sheet.setData(data);
        sheet.setClazz(data.get(0).getClass());
        sheet.setMapper(mapper);
        return sheet;
    }

    /**
     * 创建一个sheet
     *
     * @param name 名称
     * @param data 数据类型参数支持：List&lt;Object&gt;; DataRow; Map&lt;String,Object&gt;; 标准Java Bean
     * @return sheet
     * @see DataRow
     */
    public static ISheet of(String name, List<?> data) {
        return of(name, data, Collections.emptyMap());
    }

    public Class<?> getClazz() {
        return clazz;
    }

    public List<?> getData() {
        return data;
    }

    public void setData(List<?> data) {
        this.data = data;
    }

    public void setClazz(Class<?> clazz) {
        this.clazz = clazz;
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

    public String getEmptyColumn() {
        return emptyColumn;
    }

    public void setEmptyColumn(String emptyColumn) {
        this.emptyColumn = emptyColumn;
    }
}