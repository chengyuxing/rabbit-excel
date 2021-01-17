package rabbit.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import rabbit.excel.style.props.FillGround;
import rabbit.excel.style.props.Border;

import java.util.function.Consumer;

public class XStyle {
    private final CellStyle style;
    private FillGround background;
    private FillGround foreground;
    private Border border;

    public XStyle(CellStyle style) {
        this.style = style;
    }

    public void init() {
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
        return style;
    }

    public void setStyle(Consumer<CellStyle> custom) {
        custom.accept(style);
    }

    public Border getBorder() {
        return border;
    }

    public void setBorder(Border border) {
        this.border = border;
    }

    public FillGround getForeground() {
        return foreground;
    }

    public void setForeground(FillGround foreground) {
        this.foreground = foreground;
    }

    public FillGround getBackground() {
        return background;
    }

    public void setBackground(FillGround background) {
        this.background = background;
    }
}
