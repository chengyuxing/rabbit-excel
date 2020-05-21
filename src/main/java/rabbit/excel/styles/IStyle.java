package rabbit.excel.styles;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import rabbit.excel.styles.props.Background;
import rabbit.excel.styles.props.Border;
import rabbit.excel.styles.props.Foreground;

public abstract class IStyle {
    private final CellStyle style;

    public IStyle(CellStyle style) {
        this.style = style;
        init();
    }

    public CellStyle getStyle() {
        return style;
    }

    private void init() {
        Border border = border();
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
        Background background = background();
        if (background != null) {
            style.setFillBackgroundColor(background.getColor().getIndex());
            style.setFillPattern(background.getFill());
        }
        Foreground foreground = foreground();
        if (foreground != null) {
            style.setFillForegroundColor(foreground.getColor().getIndex());
            style.setFillPattern(foreground.getFill());
        }
        Font font = font();
        if (font != null) {
            style.setFont(font);
        }
    }

    public abstract Border border();

    public abstract Background background();

    public abstract Foreground foreground();

    public abstract Font font();
}
