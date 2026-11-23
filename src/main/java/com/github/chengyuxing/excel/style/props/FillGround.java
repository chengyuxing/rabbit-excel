package com.github.chengyuxing.excel.style.props;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

public class FillGround {
    private IndexedColors color;
    private FillPatternType fill;

    FillGround() {
    }

    public FillGround(IndexedColors color, FillPatternType fill) {
        this.color = color;
        this.fill = fill;
    }

    public IndexedColors getColor() {
        return color;
    }

    public FillPatternType getFill() {
        return fill;
    }
}
