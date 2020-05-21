package rabbit.excel.style.props;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

public class Border {
    private final BorderStyle borderStyle;
    private final IndexedColors borderColor;

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
