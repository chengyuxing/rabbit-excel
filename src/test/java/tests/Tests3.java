package tests;

import com.github.chengyuxing.excel.Excels;
import com.github.chengyuxing.excel.io.BigExcelLineWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.nio.file.Paths;
import java.util.UUID;

public class Tests3 {
    public static void main(String[] args) throws Exception {
    }

    @Test
    public void lineWriter() throws Exception {
        try (BigExcelLineWriter writer = Excels.bigExcelLineWriter()) {
            Sheet sheet = writer.createSheet("用户表");
            writer.writeRow(sheet, "ID", "姓名", "随机数", "UUID");
            for (int i = 0; i < 150000; i++) {
                writer.writeRow(sheet, i, "cyx", Math.random() * 1000, UUID.randomUUID());
            }

            Sheet sheet2 = writer.createSheet("空表");
            writer.writeRow(sheet2, "年龄", "家庭住址", "成员", "电话");

            writer.saveTo(Paths.get("/Users/chengyuxing/Downloads/big_excel_line_writer.xlsx"));
        }
    }

    @Test
    public void bigExcelWriter() throws Exception {
        SXSSFWorkbook wb = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
        Sheet sh = wb.createSheet();
        for (int rownum = 0; rownum < 1000000; rownum++) {
            Row row = sh.createRow(rownum);
            for (int cellnum = 0; cellnum < 10; cellnum++) {
                Cell cell = row.createCell(cellnum);
                String address = new CellReference(cell).formatAsString();
                cell.setCellValue(address);
            }
        }
        // Rows with rownum < 900 are flushed and not accessible
//        for (int rownum = 0; rownum < 900; rownum++) {
//            Assert.assertNull(sh.getRow(rownum));
//        }
        // ther last 100 rows are still in memory
//        for (int rownum = 900; rownum < 1000; rownum++) {
//            Assert.assertNotNull(sh.getRow(rownum));
//        }
        FileOutputStream out = new FileOutputStream("/Users/chengyuxing/Downloads/sxssf.xlsx");
        wb.write(out);
        out.close();
        // dispose of temporary files backing this workbook on disk
        wb.dispose();
    }
}
