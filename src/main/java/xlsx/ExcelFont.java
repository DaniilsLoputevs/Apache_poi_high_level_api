package xlsx;

import lombok.Builder;
import lombok.Data;
import org.apache.poi.xssf.usermodel.XSSFFont;

@Data
@Builder
public class ExcelFont {
    private final boolean bold;
    private final Number height;
    private final String fontName;
    
    private final XSSFFont innerFont;
    
    public XSSFFont terminate() {
        innerFont.setBold(bold);
        innerFont.setFontHeight(height.doubleValue());
        innerFont.setFontName(fontName);
        
        return innerFont;
    }
    
}
