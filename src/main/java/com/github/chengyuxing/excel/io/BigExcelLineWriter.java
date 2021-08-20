package com.github.chengyuxing.excel.io;

import com.github.chengyuxing.common.io.IOutput;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collection;
import java.util.Iterator;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * 大型Excel文件行类型数据写入器
 */
public class BigExcelLineWriter implements IOutput, AutoCloseable {
    private final AtomicInteger rowId = new AtomicInteger(0);
    private final SXSSFWorkbook workbook = new SXSSFWorkbook(1);

    public BigExcelLineWriter(boolean enableGzipTempFiles) {
        workbook.setCompressTempFiles(enableGzipTempFiles);
    }

    /**
     * 创建一个sheet
     *
     * @param name sheet名称
     * @return sheet
     */
    public Sheet createSheet(String name) {
        return workbook.createSheet(name);
    }

    /**
     * 写入一行数据到指定Sheet
     *
     * @param sheet sheet
     * @param rows  行数据
     */
    public void writeRow(Sheet sheet, Collection<Object> rows) {
        Row row = sheet.createRow(rowId.getAndIncrement());
        Iterator<Object> iterator = rows.iterator();
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
        saveTo(new FileOutputStream(path + suffix));
    }

    @Override
    public void close() throws Exception {
        workbook.close();
        workbook.dispose();
    }
}
