# 基于POI 4.0以上版进行封装的Excel读写工具
- 所有方法通过Excels静态类调用
## Example

### 准备数据

```java
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
```

### 导出excel文件（默认）

```java
Excels.writer().write(ISheet.of("SheetA", list1, javaBeanMapper),
                ISheet.of("SheetB", users, javaBeanMapper),
                ISheet.of("SheetC", list2, mapper))
                .saveTo("/Users/chengyuxing/test/excels_user000000");
```

### 导出excel文件（自定义单元格样式）

 Sheet.setCellStyle需要传入一个函数，参数1位当前行，参数2位单前列的索引类型（java bean：String，DataRow：String，Map：String，List：Integer），然后根据相应的逻辑返回一个样式。

```java
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

```

### 读取Excel文件

```java
Excels.reader(new FileInputStream("/Users/chengyuxing/test/excels_user000000.xlsx"))
                .sheetAt(1, 0, 20)
                .where((i, r) -> i >= 0)
                .where((i, r) -> !r.getString("姓名").equals("cyx"))
                .stream(row -> row)
                .forEach(System.out::println);
```

### Test

```java
@Test
public void CloseTest() throws Exception {
  List<Map<String, Object>> list = new ArrayList<>();
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

  ExcelWriter writer = Excels.writer();

  Danger danger = new Danger(writer.createCellStyle());

  ISheet<Map<String, Object>, String> iSheet = ISheet.of("sheet100", list);
  iSheet.setEmptyColumn("--");    //填充空单元格
  iSheet.setCellStyle((row, key) -> {
    //c字段大于700则添加红框
    if (key.equals("c") && (double) row.get("c") > 700) {
      return danger;
    }
    return null;
  });

  writer.write(iSheet).saveTo("/Users/chengyuxing/test/styleExcel");
}
```



### 计划
1. 可自定义表格主体样式✅
2. 可自定义表头样式