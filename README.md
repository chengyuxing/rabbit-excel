# Excel read/write utils based on POI 5.X.Y

- Static type `Excels` support returns reader and writer.

- Maven dependency (jdk1.8)

  ```xml
  <dependency>
      <groupId>com.github.chengyuxing</groupId>
      <artifactId>rabbit-excel</artifactId>
      <version>4.3.9</version>
  </dependency>
  ```

## Example

### Read Excel file

```java
String[]names=new String[]{"name","age","address"};
        try(Stream<DataRow> stream=Excels.reader(Paths.get("D:/test/test.xlsx"))
        .sheetAt(1) // specify 1st sheet
        .namedHeaderAt(0) // specify header index
        .fieldMap(names)    //data fields mapping to columns
        .stream()){
        stream.limit(10)
        .map(DataRow::toMap)
        .forEach(System.out::println);
        }
```

### Write Excel file with cell style

```java
@Test
public void CloseTest()throws Exception{
        List<Map<String, Object>>list=new ArrayList<>();
        for(int i=0;i< 10000;i++){
        Map<String, Object> row=new HashMap<>();
        row.put("a","chengyuxing");
        row.put("b",i);
        row.put("c",Math.random()*1000);
        row.put("d",LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
        row.put("e","kunming");
        row.put("f",i%3==0?"":"ok");
        list.add(row);
        }

        ExcelWriter writer=Excels.writer();

        XStyle danger=writer.createStyle();
        danger.setBorder(new Border(BorderStyle.DOUBLE,IndexedColors.RED));

        ISheet xSheet=ISheet.of("sheet100",list.stream().map(DataRow::fromMap).collect(Collectors.toList()));
        xSheet.setEmptyColumn("--");    // fill empty column
        xSheet.setCellStyle((row,key,coord)->{
        // returns danger cell style by condition
        //if (key.equals("c") && (double) row.get("c") > 700) {
        //  return danger;
        //}
        // 1st row and 5th row bind danger style
        if(coord.getX()==0||coord.getX()==5){
        return danger;
        }
        return null;
        });

        writer.write(xSheet).saveTo("/Users/chengyuxing/test/styleExcel");
        }
```
