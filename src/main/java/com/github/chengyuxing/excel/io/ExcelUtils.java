package com.github.chengyuxing.excel.io;

import org.apache.poi.ss.usermodel.Cell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ExcelUtils {
    private static final Logger log = LoggerFactory.getLogger(ExcelUtils.class);
    public static final String FORMULA_PREFIX = "{{";
    public static final String FORMULA_SUFFIX = "}}";

    public static void setCellObject(Cell cell, Object value, String other) {
        if (value == null || value.equals("")) {
            cell.setCellValue(other);
            return;
        }
        if (value instanceof String) {
            String s = ((String) value).trim();
            if (s.startsWith(FORMULA_PREFIX) && s.endsWith(FORMULA_SUFFIX)) {
                String formula = s.substring(2, s.length() - 2).trim().toUpperCase();
                try {
                    cell.setCellFormula(formula);
                } catch (Exception e) {
                    cell.setCellValue(formula);
                    log.warn("set formula error!", e);
                }
                return;
            }
        }
        cell.setCellValue(value.toString());
    }
}
