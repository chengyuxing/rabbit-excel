package rabbit.excel.type;

import org.apache.poi.ss.util.CellRangeAddress;
import rabbit.common.tuple.Triple;
import rabbit.excel.style.XStyle;

import java.util.ArrayList;
import java.util.List;

/**
 * excel复杂单元格处理类
 */
public class XRow {
    private final List<String> fields = new ArrayList<>();
    private final List<Triple<String, CellRangeAddress, XStyle>> value = new ArrayList<>();
    private boolean hasFieldMap = false;
    private int i = 0;
    private int maxRowNumber = 0;
    private int maxColumnNumber = 0;

    /**
     * 添加一个字段表头映射关系单元格
     *
     * @param field         字段
     * @param name          名称
     * @param cellAddresses 可合并单元格地址范围，标准遵循Excel坐标<br>
     *                      e.g. {@code A1:F3} (3行6列):
     *                      <blockquote>
     *                      <pre>CellRangeAddress.valueOf("A1:F3")</pre>
     *                      </blockquote>
     * @param cellStyle     单元格样式
     * @return 当前行数据
     */
    public XRow set(String field, String name, CellRangeAddress cellAddresses, XStyle cellStyle) {
        CellRangeAddress actuallyAddress;
        if (isEmpty()) {
            if (cellAddresses != null) {
                actuallyAddress = cellAddresses;
            } else {
                actuallyAddress = new CellRangeAddress(0, 0, 0, 0);
            }
        } else {
            if (cellAddresses != null) {
                actuallyAddress = cellAddresses;
            } else {
                CellRangeAddress lastAddress = value.get(value.size() - 1).getItem2();
                actuallyAddress = new CellRangeAddress(lastAddress.getFirstRow(), lastAddress.getFirstRow(), lastAddress.getLastColumn() + 1, lastAddress.getLastColumn() + 1);
            }
        }

        value.add(Triple.of(name, actuallyAddress, cellStyle));
        fields.add(field);
        if (!field.startsWith("#") && !field.endsWith("#")) {
            hasFieldMap = true;
        }
        return this;
    }

    /**
     * 添加一个字段表头映射关系单元格
     *
     * @param field         字段
     * @param name          名称
     * @param cellAddresses 可合并单元格地址范围，标准遵循Excel坐标<br>
     *                      e.g. {@code A1:F3} (3行6列):
     *                      <blockquote>
     *                      <pre>CellRangeAddress.valueOf("A1:F3")</pre>
     *                      </blockquote>
     * @return 当前行数据
     */
    public XRow set(String field, String name, CellRangeAddress cellAddresses) {
        return set(field, name, cellAddresses, null);
    }

    /**
     * 添加一个字段表头映射关系单元格
     *
     * @param field     字段
     * @param name      名称
     * @param cellStyle 单元格样式
     * @return 当前行数据
     */
    public XRow set(String field, String name, XStyle cellStyle) {
        return set(field, name, null, cellStyle);
    }

    /**
     * 添加一个字段表头映射关系单元格，单元格位置为（前一个单元格的起始行，前一个单元格的结束列往后推一格）
     *
     * @param field 字段
     * @param name  名称
     * @return 当前行数据
     */
    public XRow set(String field, String name) {
        return set(field, name, null, null);
    }

    /**
     * 添加一个简单的不映射字段的单元格
     *
     * @param name          名称
     * @param cellAddresses 可合并单元格地址范围，标准遵循Excel坐标<br>
     *                      e.g. {@code A1:F3} (3行6列):
     *                      <blockquote>
     *                      <pre>CellRangeAddress.valueOf("A1:F3")</pre>
     *                      </blockquote>
     * @param cellStyle     单元格样式
     * @return 当前行数据
     */
    public XRow add(String name, CellRangeAddress cellAddresses, XStyle cellStyle) {
        return set("#" + i++ + "#", name, cellAddresses, cellStyle);
    }

    /**
     * 添加一个简单的不映射字段的单元格
     *
     * @param name          名称
     * @param cellAddresses 可合并单元格地址范围，标准遵循Excel坐标<br>
     *                      e.g. {@code A1:F3} (3行6列):
     *                      <blockquote>
     *                      <pre>CellRangeAddress.valueOf("A1:F3")</pre>
     *                      </blockquote>
     * @return 当前行数据
     */
    public XRow add(String name, CellRangeAddress cellAddresses) {
        return add(name, cellAddresses, null);
    }

    /**
     * 添加一个简单的不映射字段的单元格，单元格位置为（前一个单元格的起始行，前一个单元格的结束列往后推一格）
     *
     * @param name      名称
     * @param cellStyle 单元格样式
     * @return 当前行数据
     */
    public XRow add(String name, XStyle cellStyle) {
        return add(name, null, cellStyle);
    }

    /**
     * 添加一个简单的不映射字段的单元格，单元格位置为（前一个单元格的起始行，前一个单元格的结束列往后推一格）
     *
     * @param name 名称
     * @return 当前行数据
     */
    public XRow add(String name) {
        return add(name, null, null);
    }

    /**
     * 判断是否为空行
     *
     * @return 是否空
     */
    public boolean isEmpty() {
        return fields.isEmpty();
    }

    /**
     * 获取映射字段的位置
     *
     * @param field 字段
     * @return 字段所在位置
     */
    public int getIndex(String field) {
        return fields.indexOf(field);
    }

    /**
     * 获取映射字段的名称
     *
     * @param field 字段
     * @return 名称
     */
    public String getName(String field) {
        int index = getIndex(field);
        return value.get(index).getItem1();
    }

    /**
     * 获取单元格地址坐标
     *
     * @param field 字段
     * @return 单元格地址坐标
     */
    public CellRangeAddress getCellAddresses(String field) {
        int index = getIndex(field);
        return value.get(index).getItem2();
    }

    /**
     * 获取单元格样式
     *
     * @param field 字段名
     * @return 单元格样式
     */
    public XStyle getStyle(String field) {
        int index = getIndex(field);
        return value.get(index).getItem3();
    }

    /**
     * 获取单元格所有字段
     *
     * @return 所有字段
     */
    public List<String> getFields() {
        return fields;
    }

    /**
     * 获取单前行的最大行号
     *
     * @return 最大行号
     */
    public int getMaxRowNumber() {
        for (String field : fields) {
            CellRangeAddress cellAddresses = getCellAddresses(field);
            int rowNumber = cellAddresses.getLastRow();
            if (rowNumber > maxRowNumber) {
                maxRowNumber = rowNumber;
            }
        }
        return maxRowNumber;
    }

    /**
     * 获取单前行的最长单元格列号
     *
     * @return 最长单元格列号
     */
    public int getMaxColumnNumber() {
        for (String field : fields) {
            CellRangeAddress cellAddresses = getCellAddresses(field);
            int columnNumber = cellAddresses.getLastColumn();
            if (columnNumber > maxColumnNumber) {
                maxColumnNumber = columnNumber;
            }
        }
        return maxColumnNumber;
    }

    /**
     * 当前行内是否有字段映射关系
     *
     * @return 是否有字段映射关系
     */
    public boolean isHasFieldMap() {
        return hasFieldMap;
    }
}
