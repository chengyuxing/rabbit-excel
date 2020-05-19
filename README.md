# 基于POI 4.0以上版进行封装的Excel读写工具
- 所有方法通过Excels静态类调用
## Example

```java
@Test
public void test1() throws IOException, InvalidFormatException {
  List<Map<String, Object>> list = new ArrayList<>();
  for (int i = 0; i < 10; i++) {
    Map<String, Object> map = new HashMap<>();
    map.put("name", "chengyuxing");
    map.put("age", i);
    map.put("address", "昆明市西山区");
    map.put("enable", i % 2 == 0);
    list.add(map);
  }

  Excels.write("/Users/chengyuxing/test/me", ISheet.ofMap("oSheet", list));

  Excels.read("/Users/chengyuxing/test/me.xlsx")
    .sheetAt(0)
    .where((idx, row) -> row.get("enable").equals("true"))
    .where((idx, row) -> row.getInt("age") > 4)
    .stream(DataRow::toMap)
    .forEach(System.out::println);
}
```

```java
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

  ISheet firstSheet = ISheet.of("Sheet1", list1);

  Excels.write("/Users/chengyuxing/test/excels_user",
               firstSheet,
               ISheet.of("Sheet10", users),
               ISheet.of("Sheet3", list2, mapper));

  Excels.read(new FileInputStream("/Users/chengyuxing/test/excels_user.xlsx"))
    .sheetAt(1, 0, 20)
    .where((i, r) -> i >= 0)
    .where((i, r) -> !r.getString("姓名").equals("cyx"))
    .stream(row -> row)
    .forEach(System.out::println);
}
```

