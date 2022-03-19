package xlsx.tools;

import org.apache.poi.ss.usermodel.Sheet;
import xlsx.core.ExcelSheetConfig;
import xlsx.utils.Pair;

import static xlsx.core.ExcelBookWriter.ABOUT_STANDARD_WIDTH_EXCEL_CHAR;


public class ExcelSheetConfigs {
    
    public static ExcelSheetConfig config() {
        return new ExcelSheetConfig();
    }
    
    /**
     * If the appearance of the report is important to you, choose the width for your font.
     * You can set the width in excel pixels using method below {@link ExcelSheetConfigs#columnWidthPixel(int, int)}.
     * See more info about column width in {@link Sheet#setColumnWidth}
     *
     * @param colIndex     -
     * @param widthInUnits the width is set in units, for different fonts, it can be a DIFFERENT that real width will bee.
     * @return option for ExcelSheetConfig.
     */
    public static Pair<Integer, Integer> columnWidth(int colIndex, int widthInUnits) {
        return new Pair<>(colIndex, (widthInUnits * ABOUT_STANDARD_WIDTH_EXCEL_CHAR));
    }
    
    public static Pair<Integer, Integer> columnWidthPixel(int colIndex, int width) {
        return new Pair<>(colIndex, width);
    }
    
}
