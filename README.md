# 基于POI 4.0以上版进行封装的Excel读写工具
- 所有方法通过Excels静态类调用

- Maven dependency (jdk1.8)

  ```xml
  <dependency>
      <groupId>com.github.chengyuxing</groupId>
      <artifactId>rabbit-excel</artifactId>
      <version>4.1.9</version>
  </dependency>
  ```
## Example

### 读取Excel文件


```java
String[] names = new String[]{"name", "age", "address"};
try (Stream<DataRow> stream = Excels.reader(Paths.get("D:/test/test.xlsx"))
        .sheetAt(1) // 指定读取第几个sheet
        .namedHeaderAt(0) // 指定表头在哪一行
        .fieldMap(names)    //翻译表头填充字段
        .stream()) {
            stream.limit(10)
                    .map(DataRow::toMap)
                    .forEach(System.out::println);
        }
```

### 导出excel文件（自定义单元格样式）

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

  XStyle danger = writer.createStyle();
  danger.setBorder(new Border(BorderStyle.DOUBLE, IndexedColors.RED));

  ISheet xSheet = ISheet.of("sheet100", list.stream().map(DataRow::fromMap).collect(Collectors.toList()));
  xSheet.setEmptyColumn("--");    //填充空单元格
  xSheet.setCellStyle((row, key, coord) -> {
    //c字段大于700则添加红框
    //if (key.equals("c") && (double) row.get("c") > 700) {
    //  return danger;
    //}
    // 第一行和第五行添加红框
    if (coord.getX() == 0 || coord.getX() == 5) {
        return danger;
    }
    return null;
  });

  writer.write(xSheet).saveTo("/Users/chengyuxing/test/styleExcel");
}
```
