package com.github.chengyuxing.excel.style;

import com.github.chengyuxing.excel.style.props.Border;
import com.github.chengyuxing.excel.style.props.FillGround;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.function.Consumer;

public class XStyle {
    private boolean built = false;
    private final CellStyle style;
    private FillGround background;
    private FillGround foreground;
    private Border border;

    public XStyle(CellStyle style) {
        this.style = style;
    }

    public void build() {
        built = true;
        if (border != null) {
            style.setBorderBottom(border.getBorderStyle());
            style.setBorderTop(border.getBorderStyle());
            style.setBorderRight(border.getBorderStyle());
            style.setBorderLeft(border.getBorderStyle());
            style.setBottomBorderColor(border.getBorderColor().getIndex());
            style.setTopBorderColor(border.getBorderColor().getIndex());
            style.setLeftBorderColor(border.getBorderColor().getIndex());
            style.setRightBorderColor(border.getBorderColor().getIndex());
        }
        if (background != null) {
            style.setFillBackgroundColor(background.getColor().getIndex());
            style.setFillPattern(background.getFill());
        }
        if (foreground != null) {
            style.setFillForegroundColor(foreground.getColor().getIndex());
            style.setFillPattern(foreground.getFill());
        }
    }

    public CellStyle getStyle() {
        if (!built) {
            build();
        }
        return style;
    }

    public void setStyle(Consumer<CellStyle> custom) {
        built = false;
        custom.accept(style);
    }

    public Border getBorder() {
        return border;
    }

    public void setBorder(Border border) {
        built = false;
        this.border = border;
    }

    public FillGround getForeground() {
        return foreground;
    }

    public void setForeground(FillGround foreground) {
        built = false;
        this.foreground = foreground;
    }

    public FillGround getBackground() {
        return background;
    }

    public void setBackground(FillGround background) {
        built = false;
        this.background = background;
    }
}
