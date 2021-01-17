package rabbit.excel.type;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel表头构建类
 */
public class XHeader {
    private final List<XRow> rows = new ArrayList<>();
    private int maxRowNumber = 0;
    private int maxColumnNumber = 0;

    /**
     * 添加一行表头
     *
     * @param row 行
     * @return 当前表头
     */
    public XHeader add(XRow row) {
        if (!isEmpty()) {
            XRow lastRow = rows.get(rows.size() - 1);
            int maxLastRow = 0;
            for (String field : lastRow.getFields()) {
                CellRangeAddress cellAddresses = lastRow.getCellAddresses(field);
                if (cellAddresses.getLastRow() > maxLastRow) {
                    maxLastRow = cellAddresses.getLastRow();
                }
            }
            List<String> currentFields = row.getFields();
            for (String currentField : currentFields) {
                CellRangeAddress cellAddresses = row.getCellAddresses(currentField);
                if (cellAddresses.getFirstRow() == 0 && cellAddresses.getLastRow() == 0) {
                    int nextRowNumber = maxLastRow + 1;
                    cellAddresses.setFirstRow(nextRowNumber);
                    cellAddresses.setLastRow(nextRowNumber);
                }
            }
        }
        rows.add(row);
        if (row.getMaxRowNumber() > maxRowNumber) {
            maxRowNumber = row.getMaxRowNumber();
        }
        if (row.getMaxColumnNumber() > maxColumnNumber) {
            maxColumnNumber = row.getMaxColumnNumber();
        }
        return this;
    }

    /**
     * 判断表头是否空
     *
     * @return 是否空
     */
    public boolean isEmpty() {
        return rows.isEmpty();
    }

    /**
     * 获取整体表头所占的最大行号
     *
     * @return 最大行号
     */
    public int getMaxRowNumber() {
        return maxRowNumber;
    }

    /**
     * 获取整体表头最长的单元格索引
     *
     * @return 最长单元格索引
     */
    public int getMaxColumnNumber() {
        return maxColumnNumber;
    }

    /**
     * 获取表头数据行
     *
     * @return 表头数据行
     */
    public List<XRow> getRows() {
        return rows;
    }
}
