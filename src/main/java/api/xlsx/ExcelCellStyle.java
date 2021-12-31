package api.xlsx;

import lombok.Builder;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;

import java.awt.Color;
import java.util.Calendar;
import java.util.Date;

/**
 * Важно знать про формат:
 * 0 - будет Числовой формат (без знаков после запятой)
 * 0.00 - будет Числовой формат (2 знака после запятой)
 * <p>
 * Далее действительны только с: {@link Calendar}, {@link Date})
 * dd.MM.yy - будет Дата формат
 * HH:ss - будет Время формат
 * dd.MM.yy HH:ss - (другие форматы, работает так же, как Время или Дата)
 */
@Builder
public class ExcelCellStyle {
    private final XSSFCellStyle cellStyleInner;
    private final XSSFDataFormat dataFormatHelper;
    private final String format;
    private final Color foregroundColor;
    private final FillPatternType fillPattern;
    private final HorizontalAlignment horizontalAlignment;
    private final VerticalAlignment verticalAlignment;
    private final ExcelFont font;
    
    private final BorderStyle borderAllSide;
    
    private final BorderStyle borderTop;
    private final BorderStyle borderBottom;
    private final BorderStyle borderLeft;
    private final BorderStyle borderRight;
    
    
    public CellStyle terminate() {
        if (format != null) cellStyleInner.setDataFormat(dataFormatHelper.getFormat(format));
        if (foregroundColor != null) cellStyleInner.setFillForegroundColor(new XSSFColor(foregroundColor));
        if (fillPattern != null) cellStyleInner.setFillPattern(fillPattern);
        if (horizontalAlignment != null) cellStyleInner.setAlignment(horizontalAlignment);
        if (verticalAlignment != null) cellStyleInner.setVerticalAlignment(verticalAlignment);
        
        if (borderAllSide != null) {
            cellStyleInner.setBorderTop(borderAllSide);
            cellStyleInner.setBorderBottom(borderAllSide);
            cellStyleInner.setBorderLeft(borderAllSide);
            cellStyleInner.setBorderRight(borderAllSide);
        } else {
            if (borderTop != null) cellStyleInner.setBorderTop(borderTop);
            if (borderBottom != null) cellStyleInner.setBorderBottom(borderBottom);
            if (borderLeft != null) cellStyleInner.setBorderLeft(borderLeft);
            if (borderRight != null) cellStyleInner.setBorderRight(borderRight);
        }
        
        if (font != null) cellStyleInner.setFont(font.terminate());
        
        return cellStyleInner;
    }
    
}
