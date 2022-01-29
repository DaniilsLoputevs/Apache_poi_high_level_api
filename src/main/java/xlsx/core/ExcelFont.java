package xlsx.core;

import lombok.Builder;
import lombok.Data;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * @author Daniils Loputevs
 */
@Data
@Builder
public class ExcelFont {
    private final boolean bold;
    private final Number height;
    private final String fontName;
    private final IndexedColors color;
    
    private final Font innerFont;
    
    public Font terminate() {
        innerFont.setBold(bold);
        if (height != null) innerFont.setFontHeightInPoints(height.shortValue());
        if (fontName != null) innerFont.setFontName(fontName);
        if (color != null) innerFont.setColor(color.index);
        
        return innerFont;
    }
    
}
