package tests;

import com.github.chengyuxing.common.DataRow;
import com.github.chengyuxing.common.io.Lines;
import com.github.chengyuxing.excel.Excels;
import com.github.chengyuxing.excel.io.ExcelWriter;
import com.github.chengyuxing.excel.style.XStyle;
import com.github.chengyuxing.excel.style.props.Border;
import com.github.chengyuxing.excel.style.props.FillGround;
import com.github.chengyuxing.excel.type.XHeader;
import com.github.chengyuxing.excel.type.XRow;
import com.github.chengyuxing.excel.type.XSheet;
import com.healthmarketscience.jackcess.Row;
import com.healthmarketscience.jackcess.Table;
import com.healthmarketscience.jackcess.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.BeforeClass;
import org.junit.Test;

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
    public void ListTest() throws Exception {
        String[] names = new String[10];
        Arrays.fill(names, "___");
        names[5] = "A";
        System.out.println(Arrays.toString(names));
    }

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
    public void arrHeader() throws Exception {
        CellRangeAddress cellAddresses = new CellRangeAddress(1, 1, 1, 3);
        System.out.println(cellAddresses.formatAsString());
        CellRangeAddress format = CellRangeAddress.valueOf("C5:C5");
        System.out.println(format.getFirstColumn());
        System.out.println(format.getLastColumn());
        System.out.println(format.getFirstRow());
        System.out.println(format.getLastRow());
    }

    @Test
    public void test1() throws Exception {
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

        ExcelWriter writer = Excels.writer();

        XStyle center = writer.createStyle();
        center.setStyle(s -> {
            s.setAlignment(HorizontalAlignment.CENTER);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
        });

        XStyle seaBlue = writer.createStyle();
        seaBlue.setForeground(new FillGround(IndexedColors.LIGHT_ORANGE, FillPatternType.SOLID_FOREGROUND));
        seaBlue.setBorder(new Border(BorderStyle.THIN, IndexedColors.GREY_25_PERCENT));
        seaBlue.setStyle(s -> {
            s.setAlignment(HorizontalAlignment.CENTER);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
        });

        XHeader headers = new XHeader();

        XRow title = new XRow();
        title.add("学生信息统计表", CellRangeAddress.valueOf("A1:C2"));

        XRow header = new XRow();
        header.set("name", "姓名", CellRangeAddress.valueOf("A3:A4"), center)
                .add("其他信息", CellRangeAddress.valueOf("B3:C3"), center);

        XRow header2 = new XRow();
        header2.set("age", "年龄", CellRangeAddress.valueOf("B4:B4"))
                .set("address", "地址", CellRangeAddress.valueOf("H4:H4"));

        headers.add(title);
        headers.add(header);
        headers.add(header2);

        System.out.println(headers.getRows());

        XSheet sheet = XSheet.of("SheetC",
                list2.stream().map(DataRow::fromMap).collect(Collectors.toList()),
                headers);
        sheet.setHeaderStyle(seaBlue);

        writer.write(sheet).saveTo("/Users/chengyuxing/Downloads/datarow2");
    }

    static final List<Map<String, Object>> list = new ArrayList<>();

    @BeforeClass
    public static void init() {
        for (int i = 0; i < 10; i++) {
            Map<String, Object> row = new HashMap<>();
            row.put("姓名", "chengyuxing");
            row.put("编号", i);
            row.put("c", Math.random() * 1000);
            row.put("d", LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
            row.put("城市", "昆明市");
            row.put("f", i % 3 == 0 ? "" : "ok");
            list.add(row);
        }
    }

    @Test
    public void writeTest() throws Exception {

        ExcelWriter writer = Excels.bigExcelWriter();

        XStyle danger = writer.createStyle();
        danger.setBorder(new Border(BorderStyle.DOUBLE, IndexedColors.RED));

        XStyle warning = writer.createStyle();
        warning.setBorder(new Border(BorderStyle.DOUBLE, IndexedColors.GREEN));

        XStyle bold = writer.createStyle();
        bold.setStyle(s -> {
            Font font = writer.createFont();
            font.setBold(true);
            font.setItalic(true);
            s.setFont(font);
        });

        XStyle center = writer.createStyle();
        center.setStyle(s -> {
            s.setAlignment(HorizontalAlignment.CENTER);
            s.setVerticalAlignment(VerticalAlignment.CENTER);
        });

        XRow xRow = new XRow();
        xRow.add("随机数据统计表", CellRangeAddress.valueOf("A1:F2"), center);
        XHeader header = new XHeader();
        XRow xRow1 = new XRow();
        xRow1.set("d", "日期时间", CellRangeAddress.valueOf("A3:A3"));
        xRow1.set("编号", "序号");
        xRow1.set("c", "分数");
        xRow1.set("城市", "所在城市");
        xRow1.set("姓名", "测试者");
        xRow1.set("f", "状态");
        header.add(xRow);
        header.add(xRow1);

        XSheet xSheet = XSheet.of("sheet100", list.stream().map(DataRow::fromMap).collect(Collectors.toList()), header);
        xSheet.setEmptyColumn("--");    //填充空单元格
        xSheet.setHeaderStyle(bold);
        xSheet.setCellStyle((row, key, coord) -> {
            //c字段大于700则添加红框
            if (key.equals("c") && (double) row.get("c") > 700) {
                return danger;
            }
            // 第一行和第五行添加绿框
            if (coord.getX() == 0 || coord.getX() == 5) {
                return warning;
            }
            return null;
        });

        writer.write(xSheet).saveTo("/Users/chengyuxing/Downloads/sxxx");
    }

    @Test
    public void readTest() throws Exception {
        String[] names = new String[]{"name", "age", "address", "1", "2", "3"};
        try (Stream<DataRow> stream = Excels.reader(Paths.get("/Users/chengyuxing/test/datarow2.xlsx"))
                .sheetAt(1) // 指定读取第几个sheet
//                .namedHeaderAt(0) // 指定表头在哪一行
                .fieldMap(names)    //翻译表头填充字段
                .stream()) {
            stream.map(d -> d.removeIf((k, v) -> v == null || v.equals("")))
                    .forEach(System.out::println);
        }
    }

    @Test
    public void tsv() throws Exception {
        try (Stream<List<String>> stream = Lines.readLines(new FileInputStream("/Users/chengyuxing/Downloads/x.tsv"), "\t")) {
            stream.limit(2)
                    .forEach(System.out::println);
        }
    }
}
