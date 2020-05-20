package rabbit.excel.types;


import rabbit.common.types.DataRow;
import rabbit.excel.styles.IStyle;

import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;

/**
 * Excel Sheet数据类
 * @param <T> 行数据类型参数
 * @param <U> 行数据索引类型（java bean：String，DataRow：String，Map：String，List：Integer）
 */
public class ISheet<T, U> {
    private String name;
    private Map<String, String> mapper;
    private List<T> data;
    private Class<T> clazz;
    private String emptyColumn = "";
    private BiFunction<T, U, IStyle> cellStyle;

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
    @SuppressWarnings("unchecked")
    public static <T, U> ISheet<T, U> of(String name, List<T> data, Map<String, String> mapper) {
        ISheet<T, U> sheet = new ISheet<>();
        sheet.setName(name);
        sheet.setData(data);
        sheet.setClazz((Class<T>) data.get(0).getClass());
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
    public static <T, U> ISheet<T, U> of(String name, List<T> data) {
        return of(name, data, Collections.emptyMap());
    }

    public BiFunction<T, U, IStyle> getCellStyle() {
        return cellStyle;
    }

    public void setCellStyleCall(BiFunction<T, U, IStyle> cellStyle) {
        this.cellStyle = cellStyle;
    }

    public Class<T> getClazz() {
        return clazz;
    }

    public List<T> getData() {
        return data;
    }

    public void setData(List<T> data) {
        this.data = data;
    }

    public void setClazz(Class<T> clazz) {
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