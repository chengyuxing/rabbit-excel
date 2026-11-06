package com.github.chengyuxing.excel.io;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

public class ExcelUtils {
    private static final Logger log = LoggerFactory.getLogger(ExcelUtils.class);
    public static final String FORMULA_PREFIX = "{{";
    public static final String FORMULA_SUFFIX = "}}";

    public static void setCellObject(Cell cell, Object value) {
        if (value == null) return;
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
            cell.setCellValue(s);
            return;
        }
        if (value instanceof Date) {
            cell.setCellValue((Date) value);
            return;
        }
        if (value instanceof Double) {
            cell.setCellValue((Double) value);
            return;
        }
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
            return;
        }
        if (value instanceof Long) {
            cell.setCellValue((Long) value);
            return;
        }
        if (value instanceof Float) {
            cell.setCellValue((Float) value);
            return;
        }
        if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
            return;
        }
        if (value instanceof Short) {
            cell.setCellValue((Short) value);
            return;
        }
        if (value instanceof Byte) {
            cell.setCellValue((Byte) value);
            return;
        }
        if (value instanceof BigDecimal) {
            cell.setCellValue(((BigDecimal) value).doubleValue());
            return;
        }
        if (value instanceof LocalDate) {
            cell.setCellValue((LocalDate) value);
            return;
        }
        if (value instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) value);
            return;
        }
        if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
            return;
        }
        if (value instanceof RichTextString) {
            cell.setCellValue((RichTextString) value);
            return;
        }
        cell.setCellValue(value.toString());
    }
}
