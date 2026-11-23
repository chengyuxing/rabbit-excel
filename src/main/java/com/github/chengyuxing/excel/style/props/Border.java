package com.github.chengyuxing.excel.style.props;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

public class Border {
    private BorderStyle borderStyle;
    private IndexedColors borderColor;

    Border() {
    }

    public Border(BorderStyle borderStyle, IndexedColors borderColor) {
        this.borderStyle = borderStyle;
        this.borderColor = borderColor;
    }

    public BorderStyle getBorderStyle() {
        return borderStyle;
    }

    public IndexedColors getBorderColor() {
        return borderColor;
    }
}
