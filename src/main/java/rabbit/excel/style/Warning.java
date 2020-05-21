package rabbit.excel.style;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import rabbit.excel.style.props.Background;
import rabbit.excel.style.props.Border;
import rabbit.excel.style.props.Foreground;

public class Warning extends IStyle {
    public Warning(CellStyle style) {
        super(style);
    }

    @Override
    public Border border() {
        return new Border(BorderStyle.THIN, IndexedColors.ORANGE);
    }

    @Override
    public Background background() {
        return null;
    }

    @Override
    public Foreground foreground() {
        return null;
    }

    @Override
    public Font font() {
        return null;
    }
}
