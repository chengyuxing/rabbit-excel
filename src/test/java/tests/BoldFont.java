package tests;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import rabbit.excel.style.IStyle;
import rabbit.excel.style.props.Background;
import rabbit.excel.style.props.Border;
import rabbit.excel.style.props.Foreground;

public class BoldFont extends IStyle {
    private final Font font;

    /**
     * 构造函数
     *
     * @param style 单元格样式
     */
    public BoldFont(CellStyle style, Font font) {
        super(style);
        this.font = font;
    }

    @Override
    public Border border() {
        return null;
    }

    @Override
    public Background background() {
        return new Background(IndexedColors.ORANGE, FillPatternType.ALT_BARS);
    }

    @Override
    public Foreground foreground() {
        return null;
    }

    @Override
    public Font font() {
        font.setBold(true);
        font.setItalic(true);
        return font;
    }
}
