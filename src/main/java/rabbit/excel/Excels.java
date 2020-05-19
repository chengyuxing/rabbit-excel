package rabbit.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import rabbit.excel.io.ExcelReader;
import rabbit.excel.io.ExcelWriter;
import rabbit.excel.types.ISheet;

import java.io.*;
import java.lang.reflect.InvocationTargetException;

/**
 * Excel文件读写操作类
 */
public final class Excels {
    public final static Logger log = LoggerFactory.getLogger(Excels.class);

    /**
     * 读Excel
     *
     * @param stream 输入流
     * @return Excel对象
     */
    public static ExcelReader read(InputStream stream) {
        return new ExcelReader(stream);
    }

    /**
     * 读Excel
     *
     * @param name 文件名
     * @return Excel
     * @throws FileNotFoundException ex
     */
    public static ExcelReader read(String name) throws FileNotFoundException {
        return read(new FileInputStream(name));
    }

    /**
     * 读Excel
     *
     * @param fileBytes 文件字节
     * @return Excel
     */
    public static ExcelReader read(byte[] fileBytes) {
        return read(new ByteArrayInputStream(fileBytes));
    }

    /**
     * 写Excel
     *
     * @param iSheet Sheet
     * @param more   更多Sheet
     * @return 字节流
     */
    public static byte[] write(ISheet iSheet, ISheet... more) {
        ISheet[] sheets = new ISheet[more.length + 1];
        sheets[0] = iSheet;
        System.arraycopy(more, 0, sheets, 1, more.length);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            for (ISheet s : sheets) {
                Sheet sheet = workbook.createSheet(s.getName());
                ExcelWriter.writeSheet(sheet, s);
            }
            workbook.write(out);
        } catch (IOException e) {
            log.error("io ex:{}", e.getMessage());
        } catch (NoSuchMethodException e) {
            log.error("no such method:{}", e.getMessage());
        } catch (InvocationTargetException e) {
            log.error("this object invoke failed:{}", e.getMessage());
        } catch (NoSuchFieldException e) {
            log.error("no such field:{}", e.getMessage());
        } catch (IllegalAccessException e) {
            log.error("access failed:{}", e.getMessage());
        }
        return out.toByteArray();
    }

    /**
     * 写Excel到输出流
     *
     * @param output 输出流
     * @param iSheet sheet
     * @param more   更多sheet
     * @throws IOException ex
     */
    public static void write(OutputStream output, ISheet iSheet, ISheet... more) throws IOException {
        output.write(write(iSheet, more));
        output.flush();
        output.close();
    }

    /**
     * 写Excel到文件
     *
     * @param path   输出路径(包括文件名)
     * @param iSheet sheet
     * @param more   更多sheet
     * @throws IOException ex
     */
    public static void write(String path, ISheet iSheet, ISheet... more) throws IOException {
        write(new FileOutputStream(fixedPath(path)), iSheet, more);
    }

    private static String fixedPath(String path) {
        if (!path.endsWith(".xlsx"))
            path += ".xlsx";
        return path;
    }
}
