package com.github.chengyuxing.excel.io;

import com.github.chengyuxing.common.io.IOutput;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Collection;
import java.util.Iterator;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Big Excel file line-mode data writer.
 */
public class BigExcelLineWriter implements IOutput, AutoCloseable {
    private final ConcurrentHashMap<String, AtomicInteger> sheetRowNumber = new ConcurrentHashMap<>();
    private final SXSSFWorkbook workbook = new SXSSFWorkbook(1);

    /**
     * Constructs a BigExcelLineWriter with enableGzipTempFiles flag.
     *
     * @param enableGzipTempFiles enable gzip
     */
    public BigExcelLineWriter(boolean enableGzipTempFiles) {
        workbook.setCompressTempFiles(enableGzipTempFiles);
    }

    /**
     * Create a sheet.
     *
     * @param name sheet name
     * @return sheet
     */
    public Sheet createSheet(String name) {
        if (sheetRowNumber.containsKey(name)) {
            throw new IllegalStateException("sheet name '" + name + "' already exists.");
        }
        sheetRowNumber.put(name, new AtomicInteger(0));
        return workbook.createSheet(name);
    }

    /**
     * Write 1 row data into sheet.
     *
     * @param sheet   sheet
     * @param rowData row data
     */
    public void writeRow(Sheet sheet, Collection<Object> rowData) {
        String sheetName = sheet.getSheetName();
        if (sheetRowNumber.containsKey(sheetName)) {
            Row row = sheet.createRow(sheetRowNumber.get(sheetName).getAndIncrement());
            Iterator<Object> iterator = rowData.iterator();
            int i = 0;
            while (iterator.hasNext()) {
                Cell cell = row.createCell(i);
                Object value = iterator.next();
                if (value == null) {
                    cell.setCellValue("");
                } else {
                    cell.setCellValue(value.toString());
                }
                i++;
            }
            return;
        }
        throw new IllegalStateException("sheet '" + sheetName + "' not exists.");
    }

    /**
     * Write 1 row data into sheet.
     *
     * @param sheet   sheet
     * @param rowData row data
     */
    public void writeRow(Sheet sheet, Object... rowData) {
        writeRow(sheet, Arrays.asList(rowData));
    }

    @Override
    public byte[] toBytes() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        return out.toByteArray();
    }

    @Override
    public void saveTo(String path) throws IOException {
        String suffix = "";
        if (!path.endsWith(".xlsx")) {
            suffix = ".xlsx";
        }
        saveTo(Files.newOutputStream(Paths.get(path + suffix)));
    }

    @Override
    public void close() throws Exception {
        workbook.close();
        workbook.dispose();
    }
}
