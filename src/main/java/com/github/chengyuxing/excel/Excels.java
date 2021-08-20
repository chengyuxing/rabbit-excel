package com.github.chengyuxing.excel;

import com.github.chengyuxing.excel.io.BigExcelLineWriter;
import com.github.chengyuxing.excel.io.ExcelReader;
import com.github.chengyuxing.excel.io.ExcelWriter;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Excel文件读写操作类
 */
public final class Excels {
    /**
     * 读Excel
     *
     * @param stream 输入流
     * @return Excel读取类
     * @throws IOException ex
     */
    public static ExcelReader reader(InputStream stream) throws IOException {
        return new ExcelReader(stream);
    }

    /**
     * 读Excel
     *
     * @param name 文件名
     * @return Excel读取类
     * @throws IOException ex
     */
    public static ExcelReader reader(String name) throws IOException {
        return reader(new FileInputStream(name));
    }

    /**
     * 读Excel
     *
     * @param path 文件名
     * @return Excel读取类
     * @throws IOException ex
     */
    public static ExcelReader reader(Path path) throws IOException {
        return reader(Files.newInputStream(path));
    }

    /**
     * 读Excel
     *
     * @param file 文件对象
     * @return Excel读取类
     * @throws IOException ex
     */
    public static ExcelReader reader(File file) throws IOException {
        return reader(new FileInputStream(file));
    }

    /**
     * 读Excel
     *
     * @param fileBytes 文件字节
     * @return Excel读取类
     * @throws IOException ex
     */
    public static ExcelReader reader(byte[] fileBytes) throws IOException {
        return reader(new ByteArrayInputStream(fileBytes));
    }

    /**
     * 获取写Excel写入器
     *
     * @return Excel写入器
     */
    public static ExcelWriter writer() {
        return new ExcelWriter(new XSSFWorkbook());
    }

    /**
     * 获取写大型Excel写入器
     *
     * @return 大型Excel写入器
     */
    public static ExcelWriter bigExcelWriter() {
        SXSSFWorkbook workbook = new SXSSFWorkbook(18);
        workbook.setCompressTempFiles(true);
        return new ExcelWriter(workbook);
    }

    /**
     * 获取按行写大型Excel写入器
     *
     * @return 大型按行Excel写入器
     */
    public static BigExcelLineWriter bigExcelLineWriter() {
        return new BigExcelLineWriter();
    }
}
