package tests;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import rabbit.excel.Excels;
import rabbit.excel.io.ExcelWriter;
import rabbit.excel.styles.Danger;
import rabbit.excel.styles.Success;
import rabbit.excel.types.ISheet;

import java.io.FileInputStream;
import java.util.*;

public class Tests {

    @Test
    public void test1() throws Exception {
        List<List<Object>> list1 = Arrays.asList(
                Arrays.asList("a", "b", "c", "d"),
                Arrays.asList("e", "f", "g", "h", "i")
        );
        List<Map<String, Object>> list2 = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put("name", "chengyuxing");
        map.put("age", 21);
        map.put("address", "kunming");

        Map<String, Object> map1 = new HashMap<>();
        map1.put("name", "chengyuxing3");
        map1.put("age", 29);
        map1.put("address", "kunming");

        Map<String, Object> map2 = new HashMap<>();
        map2.put("name", "chengyuxing2");
        map2.put("age", 22);
        map2.put("address", "kunming");

        list2.add(map);
        list2.add(map2);
        list2.add(map1);

        Map<String, String> mapper = new HashMap<>();
        mapper.put("name", "姓名");
        mapper.put("age", "年龄");
        mapper.put("address", "地址");

        List<User> users = Arrays.asList(
                new User("cyx", "昆明", "中国"),
                new User("Jackson", "美国得克萨斯州", "美国")
        );

        Map<String, String> javaBeanMapper = new HashMap<>();
        javaBeanMapper.put("name", "小名");
        javaBeanMapper.put("address", "家庭地址");
        javaBeanMapper.put("country", "所属国家");

        Excels.writer().write(ISheet.of("SheetA", list1, javaBeanMapper),
                ISheet.of("SheetB", users, javaBeanMapper),
                ISheet.of("SheetC", list2, mapper))
                .saveTo("/Users/chengyuxing/test/excels_user000000");

        Excels.read(new FileInputStream("/Users/chengyuxing/test/excels_user000000.xlsx"))
                .sheetAt(1, 0, 20)
                .where((i, r) -> i >= 0)
                .where((i, r) -> !r.getString("姓名").equals("cyx"))
                .stream(row -> row)
                .forEach(System.out::println);
    }

    @Test
    public void excelW() throws Exception {

        List<Map<String, Object>> list2 = new ArrayList<>();
        Map<String, Object> map = new LinkedHashMap<>();
        map.put("name", "chengyuxing");
        map.put("age", 21);
        map.put("address", "kunming");

        Map<String, Object> map1 = new LinkedHashMap<>();
        map1.put("name", "chengyuxing3");
        map1.put("age", 29);
        map1.put("address", "kunming");

        Map<String, Object> map2 = new LinkedHashMap<>();
        map2.put("name", "chengyuxing2");
        map2.put("age", 22);
        map2.put("address", "kunming");

        list2.add(map);
        list2.add(map2);
        list2.add(map1);


        List<List<Object>> list1 = Arrays.asList(
                Arrays.asList("a", "b", "c", "d"),
                Arrays.asList("e", "f", "g", "h", "i")
        );

        List<User> users = Arrays.asList(
                new User("cyx", "昆明", "中国"),
                new User("Jackson", "美国得克萨斯州", "美国")
        );

        ExcelWriter writer = Excels.writer();

        Danger danger = new Danger(writer.createCellStyle());
        Success success = new Success(writer.createCellStyle());

        ISheet<Map<String, Object>, String> sheet = ISheet.of("sheet1", list2);
        sheet.setCellStyle((row, field) -> {
            if ((int) row.get("age") % 2 != 0 && field.equals("age"))
                return danger;
            if (row.get("address").equals("kunming"))
                return success;
            return null;
        });

        ISheet<List<Object>, Integer> sheet1 = ISheet.of("sheet2", list1);
        sheet1.setCellStyle((row, index) -> {
            if (index == 2 && row.get(index).equals("c")) {
                return danger;
            }
            return null;
        });

        ISheet<User, String> userSheet = ISheet.of("users", users);
        userSheet.setCellStyle((u, field) -> {
            if (field.equals("name") && u.getName().equals("cyx")) {
                return danger;
            }
            return null;
        });

        writer.write(sheet, sheet1, userSheet)
                .saveTo("/Users/chengyuxing/test/writer.xlsx");

    }

    @Test
    public void CloseTest() throws Exception{
        Workbook workbook = new XSSFWorkbook();
        workbook.createSheet(",,");
        workbook.close();
        workbook.close();
    }
}
