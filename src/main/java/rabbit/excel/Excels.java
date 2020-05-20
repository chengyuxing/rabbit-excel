package rabbit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import rabbit.excel.io.ExcelReader;
import rabbit.excel.io.ExcelWriter;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

/**
 * Excel文件读写操作类
 */
public final class Excels {
    /**
     * 读Excel
     *
     * @param stream 输入流
     * @return Excel读取类
     */
    public static ExcelReader read(InputStream stream) {
        return new ExcelReader(stream);
    }

    /**
     * 读Excel
     *
     * @param name 文件名
     * @return Excel读取类
     * @throws FileNotFoundException ex
     */
    public static ExcelReader read(String name) throws FileNotFoundException {
        return read(new FileInputStream(name));
    }

    /**
     * 读Excel
     *
     * @param fileBytes 文件字节
     * @return Excel读取类
     */
    public static ExcelReader read(byte[] fileBytes) {
        return read(new ByteArrayInputStream(fileBytes));
    }

    /**
     * 写Excel数据
     * @return Excel写入类
     */
    public static ExcelWriter writer() {
        return new ExcelWriter(new XSSFWorkbook());
    }
}
