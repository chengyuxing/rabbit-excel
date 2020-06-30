package rabbit.excel.type;

import rabbit.common.types.DataRow;
import rabbit.common.types.TiFunction;
import rabbit.excel.style.IStyle;

import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * Excel Sheet数据类<br>
 */
public class ISheet {
    private String name;
    private Map<String, String> mapper;
    private List<DataRow> data;
    private String emptyColumn = "";
    private TiFunction<DataRow, String, Integer, IStyle> cellStyle;
    private IStyle headerStyle;

    ISheet() {
    }

    /**
     * 创建一个sheet<br>
     *
     * @param name   名称
     * @param data   数据
     * @param mapper 表头字段名称映射(字段名，列名)
     * @return sheet
     * @see DataRow
     */
    public static ISheet of(String name, List<DataRow> data, Map<String, String> mapper) {
        ISheet sheet = new ISheet();
        sheet.setName(name);
        sheet.setData(data);
        sheet.setMapper(mapper);
        return sheet;
    }

    /**
     * 创建一个sheet<br>
     *
     * @param name 名称
     * @param data 数据
     * @return sheet
     * @see DataRow
     */
    public static ISheet of(String name, List<DataRow> data) {
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
    public TiFunction<DataRow, String, Integer, IStyle> getCellStyle() {
        return cellStyle;
    }

    /**
     * 设置表体单元格样式函数<br><br>
     * e.g. c字段大于700则添加红框例子：
     * <blockquote>
     * <pre>
     *     Danger danger = new Danger(writer.createCellStyle());
     *     ISheet iSheet = ISheet.of("sheet1", list);
     *     iSheet.setCellStyle((row, key, index) {@code ->} {
     *         if (key.equals("c"){@code &&} (double) row.get("c") {@code >} 700) {
     *             return danger;
     *         }
     *         return null;
     *     });</pre>
     * </blockquote>
     *
     * @param cellStyle 单元格样式回调函数 {@code <数据行，列名，列序号>}
     * @see IStyle
     * @see rabbit.excel.style.impl.Danger
     */
    public void setCellStyle(TiFunction<DataRow, String, Integer, IStyle> cellStyle) {
        this.cellStyle = cellStyle;
    }

    public List<DataRow> getData() {
        return data;
    }

    void setData(List<DataRow> data) {
        this.data = data;
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