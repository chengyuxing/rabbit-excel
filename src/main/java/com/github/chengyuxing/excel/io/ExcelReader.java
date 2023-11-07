package com.github.chengyuxing.excel.io;

import com.github.chengyuxing.common.DataRow;
import com.github.chengyuxing.common.UncheckedCloseable;
import org.apache.poi.ss.usermodel.*;
import com.github.chengyuxing.excel.type.SheetInfo;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.function.Consumer;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Excel file reader.
 */
public class ExcelReader {
    private final Workbook workbook;
    private int sheetIndex = 0;
    private int headerIndex = 0;
    private boolean skipBlankHeaderCol = true;
    private String[] fields;

    /**
     * Constructs an ExcelReader with InputStream.
     *
     * @param inputStream excel file inputStream
     * @throws IOException if io error
     */
    public ExcelReader(InputStream inputStream) throws IOException {
        workbook = WorkbookFactory.create(inputStream);
    }

    /**
     * Get all sheets.
     *
     * @return list of sheets
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
     * Read sheet by index.
     *
     * @param sheetIndex sheet index
     * @return ExcelReader
     */
    public ExcelReader sheetAt(int sheetIndex) {
        this.sheetIndex = sheetIndex;
        return this;
    }

    /**
     * Specify the sheet header line index.
     *
     * @param headerIndex        header index
     * @param skipBlankHeaderCol skip blank header column or not
     * @return ExcelReader
     */
    public ExcelReader namedHeaderAt(int headerIndex, boolean skipBlankHeaderCol) {
        this.headerIndex = headerIndex;
        this.skipBlankHeaderCol = skipBlankHeaderCol;
        return this;
    }

    /**
     * Specify the sheet header line index and auto indexed blank column.
     *
     * @param headerIndex header index
     * @return ExcelReader
     */
    public ExcelReader namedHeaderAt(int headerIndex) {
        return namedHeaderAt(headerIndex, false);
    }

    /**
     * Specify the custom fields map to excel columns.
     *
     * @param fields fields
     * @return ExcelReader
     */
    public ExcelReader fieldMap(String[] fields) {
        this.fields = fields;
        return this;
    }

    /**
     * Lazy read excel to stream.<br>
     * Use {@code try-with-resource} wrap to auto close the stream while read to end.
     *
     * @return data stream
     */
    public Stream<DataRow> stream() {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        UncheckedCloseable close = UncheckedCloseable.wrap(workbook);
        Iterator<Row> iterator = sheet.rowIterator();
        // skip the no-need rows
        while (headerIndex > 0) {
            if (iterator.hasNext()) {
                iterator.next();
                headerIndex--;
            } else {
                break;
            }
        }
        boolean isCustomFieldMap = fields != null;
        // if fields customized, skip the default excel header row.
        if (isCustomFieldMap) {
            if (headerIndex >= 0) {
                if (iterator.hasNext()) {
                    iterator.next();
                }
            }
        }
        return StreamSupport.stream(new Spliterators.AbstractSpliterator<DataRow>(Long.MAX_VALUE, Spliterator.ORDERED) {
            String[] names = null;

            @Override
            public boolean tryAdvance(Consumer<? super DataRow> action) {
                if (!iterator.hasNext()) {
                    return false;
                }
                Row row = iterator.next();
                if (names == null) {
                    if (isCustomFieldMap) {
                        names = fields;
                    } else {
                        names = createDataHeader(row);
                    }
                }
                action.accept(createDataBody(names, row));
                return true;
            }
        }, false).onClose(close);
    }

    /**
     * Create data header by row.
     *
     * @param row sheet row
     * @return header columns
     */
    private String[] createDataHeader(Row row) {
        String[] names = new String[row.getLastCellNum()];
        for (int i = 0; i < names.length; i++) {
            Cell cell = row.getCell(i);
            if (skipBlankHeaderCol) {
                if (cell == null) {
                    continue;
                }
                Object v = getValue(cell);
                if (v == null) {
                    continue;
                }
                if (v.toString().trim().isEmpty()) {
                    continue;
                }
            }
            if (cell != null) {
                names[i] = getValue(cell).toString();
            } else {
                names[i] = "#" + i + "#";
            }
        }
        return names;
    }

    /**
     * Create data body.
     *
     * @param names header columns
     * @param row   row data
     * @return 1 row of data
     */
    private DataRow createDataBody(String[] names, Row row) {
        Object[] values = new Object[names.length];
        for (int x = 0, y = names.length; x < y; x++) {
            if (row.getCell(x) != null) {
                values[x] = getValue(row.getCell(x));
            } else {
                values[x] = "";
            }
        }
        return DataRow.of(names, values);
    }

    /**
     * Get cell value.
     *
     * @param cell cell
     * @return value
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

