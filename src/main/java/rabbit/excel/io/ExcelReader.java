package rabbit.excel.io;

import org.apache.poi.ss.usermodel.*;
import rabbit.common.types.DataRow;
import rabbit.common.types.UncheckedCloseable;
import rabbit.excel.type.SheetInfo;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.function.Consumer;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Excel读取类
 */
public class ExcelReader {
    private final Workbook workbook;
    private int sheetIndex = 0;

    /**
     * 构造函数
     *
     * @param inputStream 输入流
     * @throws IOException IOex
     */
    public ExcelReader(InputStream inputStream) throws IOException {
        workbook = WorkbookFactory.create(inputStream);
    }

    /**
     * 获取所有Sheet
     *
     * @return list
     */
    public List<SheetInfo> getSheets() {
        List<SheetInfo> sheets = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet.getPhysicalNumberOfRows() != 0) {
                String sheetName = sheet.getSheetName();
                sheets.add(SheetInfo.of(i, sheetName, sheet.getPhysicalNumberOfRows()));
            }
        }
        return Collections.unmodifiableList(sheets);
    }

    /**
     * 指定读取sheet
     *
     * @param sheetIndex sheet序号
     * @return Excel
     */
    public ExcelReader sheetAt(int sheetIndex) {
        this.sheetIndex = sheetIndex;
        return this;
    }

    /**
     * 惰性读取Excel装载为流，只有调用终端操作和短路操作才会真正开始执行<br>
     * 使用{@code try-with-resource}进行包裹，结束后将自动关闭输入流：
     *
     * @return 行数据流
     */
    public Stream<DataRow> stream() {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        UncheckedCloseable close = UncheckedCloseable.wrap(workbook);
        Iterator<Row> iterator = sheet.rowIterator();
        return StreamSupport.stream(new Spliterators.AbstractSpliterator<DataRow>(Long.MAX_VALUE, Spliterator.ORDERED) {
            String[] names = null;

            @Override
            public boolean tryAdvance(Consumer<? super DataRow> action) {
                if (!iterator.hasNext()) {
                    return false;
                }
                Row row = iterator.next();
                // 此处处理表头只创建一次
                if (names == null) {
                    names = createDataHeader(row);
                }
                action.accept(createDataBody(names, row));
                return true;
            }
        }, false).onClose(close);
    }

    /**
     * 创建数据表头，默认以第一行数据为表头
     *
     * @param row 数据行
     * @return 一组表头
     */
    private String[] createDataHeader(Row row) {
        String[] names = new String[row.getLastCellNum()];
        for (int i = 0; i < names.length; i++) {
            if (row.getCell(i) != null) {
                names[i] = getValue(row.getCell(i)).toString();
            }
        }
        return names;
    }

    /**
     * 创建行数据内容载体
     *
     * @param names 表头名
     * @param row   数据行
     * @return 数据行载体
     */
    private DataRow createDataBody(String[] names, Row row) {
        String[] types = new String[names.length];
        Object[] values = new Object[names.length];
        for (int x = 0, y = names.length; x < y; x++) {
            if (row.getCell(x) != null) {
                values[x] = getValue(row.getCell(x));
                types[x] = values[x].getClass().getName();
            } else {
                values[x] = "";
                types[x] = "null";
            }
        }
        return DataRow.of(names, types, values);
    }

    /**
     * 获取单元格的值
     *
     * @param cell 单元格
     * @return 值
     */
    private Object getValue(Cell cell) {
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
}

