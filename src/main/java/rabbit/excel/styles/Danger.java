package rabbit.excel.styles;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import rabbit.excel.styles.props.Background;
import rabbit.excel.styles.props.Border;
import rabbit.excel.styles.props.Foreground;

public class Danger extends IStyle {

    public Danger(CellStyle style) {
        super(style);
    }

    @Override
    public Border border() {
        return new Border(BorderStyle.THIN, IndexedColors.RED);
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
