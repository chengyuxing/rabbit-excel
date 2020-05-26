package rabbit.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import rabbit.excel.style.props.Background;
import rabbit.excel.style.props.Border;
import rabbit.excel.style.props.Foreground;

/**
 * 单元格样式抽象类
 */
public abstract class IStyle {
    private final CellStyle style;

    /**
     * 构造函数
     *
     * @param style 单元格样式
     */
    public IStyle(CellStyle style) {
        this.style = style;
        init();
    }

    /**
     * 获取样式
     *
     * @return 样式
     */
    public CellStyle getStyle() {
        return style;
    }

    /**
     * 初始化
     */
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

    /**
     * 边框
     *
     * @return 边框对象
     */
    public abstract Border border();

    /**
     * 背景
     *
     * @return 背景对象
     */
    public abstract Background background();

    /**
     * 前景
     *
     * @return 前景对象
     */
    public abstract Foreground foreground();

    /**
     * 字形
     *
     * @return 字形对象
     */
    public abstract Font font();
}
