package rabbit.excel.io;

import org.apache.poi.ss.usermodel.*;
import rabbit.common.types.DataRow;
import rabbit.excel.type.SheetMetaData;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.function.BiPredicate;
import java.util.function.Function;
import java.util.stream.Stream;

/**
 * Excel读取类
 */
public class ExcelReader implements AutoCloseable {
    private final InputStream inputStream;
    private final List<BiPredicate<Integer, DataRow>> filters = new ArrayList<>();
    private Workbook workbook;
    private int sheetIndex = 0;
    private int rowStart = 0;
    private int count = -1;

    /**
     * 构造函数
     *
     * @param inputStream 输入流
     */
    public ExcelReader(InputStream inputStream) {
        this.inputStream = inputStream;
    }

    /**
     * 获取所有Sheet
     *
     * @return list
     * @throws IOException e
     */
    public List<SheetMetaData> getSheets() throws IOException {
        GenWorkbookIfNecessary();
        List<SheetMetaData> sheets = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet.getPhysicalNumberOfRows() != 0) {
                String sheetName = sheet.getSheetName();
                sheets.add(SheetMetaData.of(i, sheetName, sheet.getPhysicalNumberOfRows()));
            }
        }
        return Collections.unmodifiableList(sheets);
    }

    /**
     * where条件过滤
     *
     * @param rowFilter（当前序号，当前行）
     * @return Excel
     */
    public ExcelReader where(BiPredicate<Integer, DataRow> rowFilter) {
        filters.add(rowFilter);
        return this;
    }

    /**
     * 指定读取sheet
     *
     * @param sheetIndex 序号
     * @param rowStart   开始行 从0开始
     * @param count      条数
     * @return Excel
     */
    public ExcelReader sheetAt(int sheetIndex, int rowStart, int count) {
        this.sheetIndex = sheetIndex;
        this.rowStart = rowStart;
        this.count = count;
        return this;
    }

    /**
     * 指定读取sheet
     *
     * @param sheetIndex sheet序号
     * @param rowStart   开始行 从0开始
     * @return Excel
     */
    public ExcelReader sheetAt(int sheetIndex, int rowStart) {
        return sheetAt(sheetIndex, rowStart, -1);
    }

    /**
     * 指定读取sheet
     *
     * @param sheetIndex sheet序号
     * @return Excel
     */
    public ExcelReader sheetAt(int sheetIndex) {
        return sheetAt(sheetIndex, 0, -1);
    }

    /**
     * 读取Excel装载为流
     *
     * @param convert 行数据转换
     * @param <R>     结果类型参数
     * @return 行数据流
     * @throws IOException ex
     */
    public <R> Stream<R> stream(Function<DataRow, R> convert) throws IOException {
        GenWorkbookIfNecessary();
        Stream.Builder<R> builder = Stream.builder();
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        int rowCount = sheet.getPhysicalNumberOfRows();
        if (rowCount > 0) {
            Row headerRow = sheet.getRow(0);
            String[] header = new String[headerRow.getLastCellNum()];

            for (int x = 0; x < header.length; x++) {
                if (headerRow.getCell(x) != null) {
                    header[x] = getValue(headerRow.getCell(x)).toString().toLowerCase();
                }
            }
            if (count < 1 || count > rowCount) {
                count = rowCount;
            }
            for (int i = rowStart; i < count; i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Object[] value = new Object[header.length];
                    String[] types = new String[header.length];
                    for (int x = 0, y = header.length; x < y; x++) {
                        if (row.getCell(x) != null) {
                            value[x] = getValue(row.getCell(x));
                            types[x] = value[x].getClass().getName();
                        } else {
                            value[x] = "";
                            types[x] = "null";
                        }
                    }
                    DataRow dataRow = DataRow.of(header, types, value);
                    boolean passed = true;
                    for (BiPredicate<Integer, DataRow> filter : filters) {
                        if (!filter.test(i, dataRow)) {
                            passed = false;
                            break;
                        }
                    }
                    if (passed)
                        builder.accept(convert.apply(dataRow));
                }
            }
        }
        return builder.build();
    }

    /**
     * 如果有必要就创建一个新的工作簿
     *
     * @throws IOException e
     */
    private void GenWorkbookIfNecessary() throws IOException {
        if (workbook == null)
            workbook = WorkbookFactory.create(inputStream);
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

    /**
     * 关闭工作簿
     *
     * @throws Exception 关闭异常
     */
    @Override
    public void close() throws Exception {
        workbook.close();
        filters.clear();
    }
}

