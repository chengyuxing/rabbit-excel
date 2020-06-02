package rabbit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import rabbit.excel.io.ExcelReader;
import rabbit.excel.io.ExcelWriter;

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
     * @param fileBytes 文件字节
     * @return Excel读取类
     * @throws IOException ex
     */
    public static ExcelReader reader(byte[] fileBytes) throws IOException {
        return reader(new ByteArrayInputStream(fileBytes));
    }

    /**
     * 写Excel数据
     *
     * @return Excel写入类
     */
    public static ExcelWriter writer() {
        return new ExcelWriter(new XSSFWorkbook());
    }
}
