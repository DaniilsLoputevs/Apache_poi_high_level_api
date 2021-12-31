package xlsx.tools;

import xlsx.core.ExcelBook;
import xlsx.core.ExcelCellStyle;

/**
 * @author Daniils Loputevs
 */
public class ExcelCellStyles {
    public static final ExcelCellStyle EMPTY = ExcelCellStyle.builder().build();
    public static final ExcelCellStyle DEFAULT = ExcelCellStyle.builder().build();
    
    public static ExcelCellStyle buildIdStyle(ExcelBook excelBook) {
        return excelBook.makeStyle().format("0").build();
    }
    
    public static ExcelCellStyle buildCurrencyStyle(ExcelBook excelBook) {
        return excelBook.makeStyle().format("0.00").build();
    }
}
