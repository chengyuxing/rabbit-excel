package rabbit.excel.style.impl;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import rabbit.excel.style.IStyle;
import rabbit.excel.style.props.Background;
import rabbit.excel.style.props.Border;
import rabbit.excel.style.props.Foreground;

public class SeaBlue extends IStyle {
    public SeaBlue(CellStyle style) {
        super(style);
    }

    @Override
    public Border border() {
        return null;
    }

    @Override
    public Background background() {
        return null;
    }

    @Override
    public Foreground foreground() {
        return new Foreground(IndexedColors.SEA_GREEN, FillPatternType.SOLID_FOREGROUND);
    }

    @Override
    public Font font() {
        return null;
    }
}