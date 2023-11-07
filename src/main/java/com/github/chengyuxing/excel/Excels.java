package com.github.chengyuxing.excel;

import com.github.chengyuxing.excel.io.BigExcelLineWriter;
import com.github.chengyuxing.excel.io.ExcelReader;
import com.github.chengyuxing.excel.io.ExcelWriter;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * Excel file read/write utils.
 */
public final class Excels {
    /**
     * Returns an ExcelReader with InputStream.
     *
     * @param stream excel file inputStream
     * @return ExcelReader
     * @throws IOException ex
     */
    public static ExcelReader reader(InputStream stream) throws IOException {
        return new ExcelReader(stream);
    }

    /**
     * Returns an ExcelReader with full file name.
     *
     * @param name file name
     * @return ExcelReader
     * @throws IOException ex
     */
    public static ExcelReader reader(String name) throws IOException {
        return reader(Files.newInputStream(Paths.get(name)));
    }

    /**
     * Returns an ExcelReader with Path.
     *
     * @param path file path
     * @return ExcelReader
     * @throws IOException ex
     */
    public static ExcelReader reader(Path path) throws IOException {
        return reader(Files.newInputStream(path));
    }

    /**
     * Returns an ExcelReader with File.
     *
     * @param file file
     * @return ExcelReader
     * @throws IOException ex
     */
    public static ExcelReader reader(File file) throws IOException {
        return reader(Files.newInputStream(file.toPath()));
    }

    /**
     * Returns an ExcelReader with bytes.
     *
     * @param fileBytes file bytes
     * @return ExcelReader
     * @throws IOException ex
     */
    public static ExcelReader reader(byte[] fileBytes) throws IOException {
        return reader(new ByteArrayInputStream(fileBytes));
    }

    /**
     * Returns an ExcelWriter.
     *
     * @return ExcelWriter
     */
    public static ExcelWriter writer() {
        return new ExcelWriter(new XSSFWorkbook());
    }

    /**
     * Returns a big ExcelWriter.
     *
     * @return big ExcelWriter
     */
    public static ExcelWriter bigExcelWriter() {
        SXSSFWorkbook workbook = new SXSSFWorkbook(18);
        return new ExcelWriter(workbook);
    }

    /**
     * Returns a big Excel line-mode Writer.
     *
     * @return big Excel line-mode Writer
     */
    public static BigExcelLineWriter bigExcelLineWriter() {
        return new BigExcelLineWriter(false);
    }
}
