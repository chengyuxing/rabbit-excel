package tests;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import rabbit.excel.style.Danger;
import rabbit.excel.style.Warning;

import java.io.FileOutputStream;

public class StyleTest {
    @Test
    public void testStyle() throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("style");

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("11111111");
        CellStyle style = workbook.createCellStyle();
        cell.setCellStyle(new Warning(style).getStyle());

        Cell cell2 = row.createCell(1);
        cell2.setCellValue("2222222");
        CellStyle style1 = workbook.createCellStyle();
        cell2.setCellStyle(new Danger(style1).getStyle());

        Cell cell3 = row.createCell(2);
        cell3.setCellValue("43333333");

        workbook.write(new FileOutputStream("/Users/chengyuxing/test/style.xlsx"));
    }
}
