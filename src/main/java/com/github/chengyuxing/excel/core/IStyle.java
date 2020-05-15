package com.github.chengyuxing.excel.core;

import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.rabbit.common.tuple.Pair;
import org.rabbit.common.tuple.Tuples;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Objects;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * 单元格样式
 */
public class IStyle {
    private final static Logger log = LoggerFactory.getLogger(IStyle.class);
    private static final StylesTable STYLES_TABLE = new StylesTable();
    private static List<Pair<Method, Method>> GET_SET_METHODS;
    private XSSFCellStyle globalStyle;
    private XSSFCellStyle headerStyle;
    private XSSFCellStyle bodyStyle;

    public IStyle() {

    }

    public static XSSFCellStyle bind(XSSFCellStyle source, XSSFCellStyle target) {
        if (GET_SET_METHODS == null) {
            GET_SET_METHODS = Stream.of(target.getClass().getDeclaredMethods())
                    .filter(m -> m.getName().startsWith("get"))
                    .map(Method::getName)
                    .map(m -> {
                        try {
                            Method getMethod = target.getClass().getDeclaredMethod(m);
                            Class<?> returnType = getMethod.getReturnType();
                            Method setMethod = target.getClass().getDeclaredMethod("set" + m.substring(3), returnType);
                            return Tuples.pair(getMethod, setMethod);
                        } catch (NoSuchMethodException e) {
                            return null;
                        }
                    })
                    .filter(Objects::nonNull)
                    .collect(Collectors.toList());
        }
        for (Pair<Method, Method> p : GET_SET_METHODS) {
            Method set = p.getItem2();
            Method get = p.getItem1();
            try {
                set.invoke(target, get.invoke(source));
            } catch (IllegalAccessException | InvocationTargetException e) {
                log.error("binding value error:{}", e.getMessage());
            }
        }
        return target;
    }

    public static IStyle create() {
        return new IStyle();
    }

    public IStyle globalStyle(Function<XSSFCellStyle, XSSFCellStyle> styleSetter) {
        if (globalStyle == null)
            globalStyle = STYLES_TABLE.createCellStyle();
        globalStyle = styleSetter.apply(globalStyle);
        return this;
    }

    public IStyle headerStyle(Function<XSSFCellStyle, XSSFCellStyle> styleSetter) {
        if (headerStyle == null)
            headerStyle = STYLES_TABLE.createCellStyle();
        headerStyle = styleSetter.apply(headerStyle);
        return this;
    }

    public IStyle bodyStyle(Function<XSSFCellStyle, XSSFCellStyle> styleSetter) {
        if (bodyStyle == null)
            bodyStyle = STYLES_TABLE.createCellStyle();
        bodyStyle = styleSetter.apply(bodyStyle);
        return this;
    }

    public void setBodyStyle(XSSFCellStyle bodyStyle) {
        this.bodyStyle = bodyStyle;
    }

    public void setGlobalStyle(XSSFCellStyle globalStyle) {
        this.globalStyle = globalStyle;
    }

    public void setHeaderStyle(XSSFCellStyle headerStyle) {
        this.headerStyle = headerStyle;
    }

    public XSSFCellStyle getBodyStyle() {
        return bodyStyle;
    }

    public XSSFCellStyle getGlobalStyle() {
        return globalStyle;
    }

    public XSSFCellStyle getHeaderStyle() {
        return headerStyle;
    }
}
