package rabbit.excel.styles.props;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

public class Background {
    private final IndexedColors color;
    private final FillPatternType fill;

    public IndexedColors getColor() {
        return color;
    }

    public FillPatternType getFill() {
        return fill;
    }

    public Background(IndexedColors color, FillPatternType fill) {
        this.color = color;
        this.fill = fill;
    }
}
