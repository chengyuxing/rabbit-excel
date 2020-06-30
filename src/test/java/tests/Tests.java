package tests;

import com.healthmarketscience.jackcess.*;
import org.junit.BeforeClass;
import org.junit.Test;
import rabbit.common.io.TSVReader;
import rabbit.common.io.TSVWriter;
import rabbit.common.types.DataRow;
import rabbit.excel.Excels;
import rabbit.excel.io.ExcelWriter;
import rabbit.excel.style.impl.Danger;
import rabbit.excel.style.impl.SeaBlue;
import rabbit.excel.style.impl.Success;
import rabbit.excel.type.ISheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class Tests {

    @Test
    public void AccessTest() throws IOException {
        Database db = DatabaseBuilder.open(new File("/Users/chengyuxing/test/my.mdb"));
        Table table = db.getTable("user");
        List<String> fields = table.getColumns().stream().map(Column::getName).collect(Collectors.toList());
        Row row;
        while ((row = table.getNextRow()) != null) {
            for (String field : fields) {
                System.out.print(row.get(field) + ",");
            }
            System.out.println();
        }
    }

    @Test
    public void createAccessTable() throws Exception {
        Database db = DatabaseBuilder.create(Database.FileFormat.V2016, new File("/Users/chengyuxing/test/my.mdb"));
        Table table = new TableBuilder("user")
                .addColumn(new ColumnBuilder("id").setType(DataType.LONG).setAutoNumber(true))
                .addColumn(new ColumnBuilder("name").setType(DataType.TEXT))
                .addColumn(new ColumnBuilder("address").setType(DataType.TEXT))
                .toTable(db);
        for (int i = 0; i < 100; i++) {
            table.addRow(Column.AUTO_NUMBER, "cyx" + i, "云南省昆明市");
        }
    }

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

        List<DataRow> row1 = list1.stream().map(l -> DataRow.fromList(l, "A_A", "B", "C", "D")).collect(Collectors.toList());

        Excels.writer().write(
                ISheet.of("SheetB",
                        row1),
                ISheet.of("SheetC",
                        list2.stream().map(DataRow::fromMap).collect(Collectors.toList()),
                        mapper))
                .saveTo("/Users/chengyuxing/test/datarow2");

        Excels.reader(new FileInputStream("/Users/chengyuxing/test/datarow2.xlsx"))
                .sheetAt(1)
                .stream()
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

//        writer.write(sheet, sheet1, userSheet)
//                .saveTo("/Users/chengyuxing/test/writer.xlsx");

    }

    static final List<Map<String, Object>> list = new ArrayList<>();

    @BeforeClass
    public static void init() {
        for (int i = 0; i < 10000; i++) {
            Map<String, Object> row = new HashMap<>();
            row.put("a", "chengyuxing");
            row.put("b", i);
            row.put("c", Math.random() * 1000);
            row.put("d", LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
            row.put("e", "昆明市");
            row.put("f", i % 3 == 0 ? "" : "ok");
            list.add(row);
        }
    }

    @Test
    public void writeTest() throws Exception {

        ExcelWriter writer = Excels.writer();

        Danger danger = new Danger(writer.createCellStyle());
        SeaBlue seaBlue = new SeaBlue(writer.createCellStyle());

        ISheet iSheet = ISheet.of("sheet100", list.stream().map(DataRow::fromMap).collect(Collectors.toList()));
        iSheet.setEmptyColumn("--");    //填充空单元格
        iSheet.setHeaderStyle(seaBlue);
        iSheet.setCellStyle((row, key, index) -> {
            //c字段大于700则添加红框
            if (key.equals("c") && (double) row.get("c") > 700) {
                return danger;
            }
            return null;
        });

        writer.write(iSheet).saveTo("/Users/chengyuxing/test/data_row");
    }

    @Test
    public void readTest() throws Exception {
        try (Stream<DataRow> stream = Excels.reader(Paths.get("/Users/chengyuxing/test/styleExcel.xlsx")).stream()) {
            stream.limit(1000)
                    .map(DataRow::toMap)
                    .forEach(System.out::println)
            ;
        }
    }

    @Test
    public void tsv() throws Exception {
        TSVReader tsvReader = TSVReader.of(new FileInputStream("/Users/chengyuxing/Downloads/x.tsv"));
        try (Stream<DataRow> stream = tsvReader.stream()) {
            stream//.limit(2)
                    .map(DataRow::toMap)
                    .forEach(System.out::println);
        }
    }
}
