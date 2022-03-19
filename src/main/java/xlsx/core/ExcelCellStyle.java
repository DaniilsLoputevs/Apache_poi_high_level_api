package xlsx.core;

import lombok.Builder;
import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.awt.Color;
import java.util.Date;

/**
 * ### format ###:
 * value display format wil be work for:
 * <p>
 * {@link Number}<p>
 * 0 - (no decimal signs)<p>
 * 0.00 - (2 decimal signs)<p>
 * <p>
 * Any date value (Example: {@link java.time.LocalDateTime}, {@link Date}, {@link java.time.LocalDate} and etc)<p>
 * dd.MM.yy - Date format (excel Date)<p>
 * HH:ss - Date format (excel Time)<p>
 * dd.MM.yy HH:ss - Date and Time (excel all formats)<p>
 *
 * @author Daniils Loputevs
 */
@Builder
@Data
public class ExcelCellStyle {
    private final DataFormat dataFormatHelper;
    private final String format;
    private final Color foregroundColor;
    private final IndexedColors foregroundColorIndex;
    private final FillPatternType fillPattern;
    
    private final HorizontalAlignment allSideAlignment;
    private final HorizontalAlignment horizontalAlignment;
    private final VerticalAlignment verticalAlignment;
    
    private final ExcelFont font;
    
    private final BorderStyle borderAllSide;
    private final BorderStyle borderTop;
    private final BorderStyle borderBottom;
    private final BorderStyle borderLeft;
    private final BorderStyle borderRight;
    
    CellStyle cellStyleInner;
    boolean isTerminated;
    
    
    public CellStyle terminate(CellStyle cellStyleInner) {
        if (isTerminated) return this.cellStyleInner;
        else this.cellStyleInner = cellStyleInner;
        
        if (format != null) cellStyleInner.setDataFormat(dataFormatHelper.getFormat(format));
        if (foregroundColor != null)
            if (cellStyleInner instanceof HSSFCellStyle)
                throw new IllegalStateException("java.awt.Color is support only for XSSFCellStyle OR SXSSFCellStyle.");
            else ((XSSFCellStyle) cellStyleInner).setFillForegroundColor(new XSSFColor(foregroundColor, null));
        if (foregroundColorIndex != null) cellStyleInner.setFillForegroundColor(foregroundColorIndex.index);
        if (fillPattern != null) cellStyleInner.setFillPattern(fillPattern);
        
        if (allSideAlignment != null) {
            cellStyleInner.setAlignment(allSideAlignment);
            cellStyleInner.setVerticalAlignment(VerticalAlignment.valueOf(allSideAlignment.name()));
        } else {
            if (horizontalAlignment != null) cellStyleInner.setAlignment(horizontalAlignment);
            if (verticalAlignment != null) cellStyleInner.setVerticalAlignment(verticalAlignment);
        }
        
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
        
        isTerminated = true;
        return cellStyleInner;
    }
    
}
