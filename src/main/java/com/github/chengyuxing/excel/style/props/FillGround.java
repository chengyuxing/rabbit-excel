package com.github.chengyuxing.excel.style.props;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

public class FillGround {
    private final IndexedColors color;
    private final FillPatternType fill;

    public IndexedColors getColor() {
        return color;
    }

    public FillPatternType getFill() {
        return fill;
    }

    public FillGround(IndexedColors color, FillPatternType fill) {
        this.color = color;
        this.fill = fill;
    }
}
