package rabbit.excel.type;

import rabbit.common.types.DataRow;
import rabbit.excel.style.IStyle;

import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;

/**
 * Excel Sheet数据类<br>
 *
 * @param <T> 行数据类型参数
 * @param <U> 行数据索引类型
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
     * 创建一个sheet<br>
     * 数据类型为{@code List<List<Object>>}时索引类型为{@code Integer}：
     * <blockquote>
     * <pre>{@code ISheet<List<Object>, Integer> sheet1 = ISheet.of("sheet1", list);}</pre>
     * </blockquote>
     * 数据类型为 {@link DataRow}, {@link Map}, javaBean时索引类型为{@code String}：
     * <blockquote>
     * <pre>{@code ISheet<Map<String, Object>, String> sheet = ISheet.of("sheet1", listMap);}</pre>
     * </blockquote>
     *
     * @param name   名称
     * @param data   数据
     * @param mapper 表头字段名称映射(字段名，列名)
     * @param <T>    行数据类型参数：{@code List<Object>, Map<String, Object>}, {@link DataRow}, 标准Java Bean(需指定注解Head，用于java bean的字段名映射)
     * @param <U>    索引类型参数
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
     * 创建一个sheet<br>
     * 数据类型为{@code List<List<Object>>}时索引类型为{@code Integer}：
     * <blockquote>
     * <pre>{@code ISheet<List<Object>, Integer> sheet1 = ISheet.of("sheet1", list);}</pre>
     * </blockquote>
     * 数据类型为 {@link DataRow}, {@link Map}, javaBean时索引类型为{@code String}：
     * <blockquote>
     * <pre>{@code ISheet<Map<String, Object>, String> sheet = ISheet.of("sheet1", listMap);}</pre>
     * </blockquote>
     *
     * @param name 名称
     * @param data 数据
     * @param <T>  行数据类型参数：{@code List<Object>, Map<String, Object>}, {@link DataRow}, 标准Java Bean(需指定注解Head，用于java bean的字段名映射)
     * @param <U>  索引类型参数
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
     * @see rabbit.excel.style.impl.Danger
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
     * 设置表体单元格样式函数<br><br>
     * c字段大于700则添加红框例子：
     * <blockquote>
     * <pre>
     *     Danger danger = new Danger(writer.createCellStyle());
     *     ISheet{@code <Map<String, Object>, String>} iSheet = ISheet.of("sheet1", list);
     *     iSheet.setCellStyle((row, key) {@code ->} {
     *         if (key.equals("c") && (double) row.get("c") {@code >} 700) {
     *             return danger;
     *         }
     *         return null;
     *     });</pre>
     * </blockquote>
     *
     * @param cellStyle 单元格样式回调函数
     * @see IStyle
     * @see rabbit.excel.style.impl.Danger
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