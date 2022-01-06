package xlsx.core;

import lombok.Builder;
import lombok.Data;
import org.apache.poi.ss.usermodel.Font;

/**
 * @author Daniils Loputevs
 */
@Data
@Builder
public class ExcelFont {
    private final boolean bold;
    private final Number height;
    private final String fontName;
    
    private final Font innerFont;
    
    public Font terminate() {
        innerFont.setBold(bold);
        if (height != null) innerFont.setFontHeightInPoints(height.shortValue());
        if (fontName != null) innerFont.setFontName(fontName);
        
        return innerFont;
    }
    
}
