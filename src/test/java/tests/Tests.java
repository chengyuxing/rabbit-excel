package tests;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.junit.Test;
import rabbit.excel.Excels;
import rabbit.excel.types.ISheet;

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

        Excels.write("/Users/chengyuxing/test/excels_user",
                ISheet.of("SheetA", list1,javaBeanMapper),
                ISheet.of("SheetB", users, javaBeanMapper),
                ISheet.of("SheetC", list2, mapper));

//        Excels.read(new FileInputStream("/Users/chengyuxing/test/excels_user.xlsx"))
//                .sheetAt(1, 0, 20)
//                .where((i, r) -> i >= 0)
//                .where((i, r) -> !r.getString("姓名").equals("cyx"))
//                .stream(row -> row)
//                .forEach(System.out::println);
    }

    @Test
    public void methods() throws Exception {
        XSSFCellStyle source = new StylesTable().createCellStyle();
        XSSFCellStyle target = new StylesTable().createCellStyle();

        source.setAlignment(HorizontalAlignment.CENTER);
        source.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());

        System.out.println(target.getAlignment());
        System.out.println(target.getFillBackgroundColor());
    }
}
