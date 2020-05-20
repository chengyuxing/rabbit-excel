package rabbit.excel.types;

import rabbit.common.types.DataRow;
import rabbit.excel.styles.IStyle;

import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;

/**
 * Excel Sheet数据类
 *
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
    private IStyle headerStyle;

    ISheet() {
    }

    /**
     * 创建一个sheet
     *
     * @param name   名称
     * @param data   数据
     * @param mapper 表头字段名称映射(字段名，列名)
     * @param <T>    行数据类型参数：List&lt;Object&gt;; DataRow; Map&lt;String,Object&gt;; 标准Java Bean(需指定注解Head，用于java bean的字段名映射)
     * @param <U>    索引类型参数，除行数据类型List&lt;Object&gt;索引为Integer类型(单元格序号)以外，其他都为String类型(字段名)
     * @return sheet
     * @see DataRow
     * @see Head
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
     * @param data 数据
     * @param <T>  行数据类型参数：List&lt;Object&gt;; DataRow; Map&lt;String,Object&gt;; 标准Java Bean(需指定注解Head，用于java bean的字段名映射)
     * @param <U>  索引类型参数，除行数据类型List&lt;Object&gt;索引为Integer类型(单元格序号)以外，其他都为String类型(字段名)
     * @return sheet
     * @see DataRow
     * @see Head
     */
    public static <T, U> ISheet<T, U> of(String name, List<T> data) {
        return of(name, data, Collections.emptyMap());
    }

    /**
     * 获取表头样式
     *
     * @return 表头样式
     */
    public IStyle getHeaderStyle() {
        return headerStyle;
    }

    /**
     * 设置表头样式
     *
     * @param headerStyle 表头样式
     */
    public void setHeaderStyle(IStyle headerStyle) {
        this.headerStyle = headerStyle;
    }

    /**
     * 获取表体单元格样式函数
     *
     * @return 表体单元格样式函数
     */
    public BiFunction<T, U, IStyle> getCellStyle() {
        return cellStyle;
    }

    /**
     * 设置表体单元格样式函数
     */
    public void setCellStyle(BiFunction<T, U, IStyle> cellStyle) {
        this.cellStyle = cellStyle;
    }

    public Class<T> getClazz() {
        return clazz;
    }

    public List<T> getData() {
        return data;
    }

    void setData(List<T> data) {
        this.data = data;
    }

    void setClazz(Class<T> clazz) {
        this.clazz = clazz;
    }

    public String getName() {
        return name;
    }

    void setName(String name) {
        this.name = name;
    }

    public Map<String, String> getMapper() {
        return mapper;
    }

    void setMapper(Map<String, String> mapper) {
        this.mapper = mapper;
    }

    public String getEmptyColumn() {
        return emptyColumn;
    }

    public void setEmptyColumn(String emptyColumn) {
        this.emptyColumn = emptyColumn;
    }
}