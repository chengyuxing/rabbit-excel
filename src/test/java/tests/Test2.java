package tests;

import com.github.chengyuxing.common.DataRow;
import org.junit.BeforeClass;
import org.junit.Test;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class Test2 {
    private static final List<Map<String, Object>> list = new ArrayList<>();

    @BeforeClass
    public static void init() {
        for (int i = 0; i < 100000; i++) {
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
    public void toMap() {
        List<DataRow> rows = list.stream()
                .map(DataRow::fromMap)
                .collect(Collectors.toList());
        System.out.println(rows.size());
    }

    @Test
    public void toDataRow() {
    }
}
