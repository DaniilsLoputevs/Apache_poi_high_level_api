package xlsx.tools;

import lombok.val;
import org.apache.poi.ss.util.CellRangeAddress;
import xlsx.core.*;

import java.util.List;
import java.util.function.Consumer;

/**
 * @author Daniils Loputevs
 */
public final class ExcelCellGroupSelectors {
    
    public static ExcelCellGroupSelector cellGroupSelector(ExcelCellGroupType type, String collectPattern) {
        return new ExcelCellGroupSelector(type, collectPattern);
    }
    
    public static Consumer<List<ExcelCell>> setValueAndHeaderForGroup(String value, ExcelCellStyle headerStyle) {
        return (List<ExcelCell> cells) -> {
            if (cells.isEmpty()) return;
            for (val cell : cells) {
                cell.setValue(value);
                cell.setStyle(headerStyle);
            }
        };
    }
    
    public static Consumer<List<ExcelCell>> mergeCellGroupAndSetValueAndStyle(String value, ExcelCellStyle style, ExcelBook book) {
        return (List<ExcelCell> cells) -> {
            if (cells.size() == 1) throw new IllegalStateException("try to merge 1 cell, not many cells");
            int rowStartIndex = Integer.MAX_VALUE, rowEndIndex = 0, colStartIndex = Integer.MAX_VALUE, colEndIndex = 0;
            for (val cell : cells) {
                cell.setValue(value);
                cell.setStyle(style);
                rowStartIndex = Math.min(rowStartIndex, cell.getRowIndex());
                rowEndIndex = Math.max(rowEndIndex, cell.getRowIndex());
                
                colStartIndex = Math.min(colStartIndex, cell.getColIndex());
                colEndIndex = Math.max(colEndIndex, cell.getColIndex());
            }
            book.getFirstWorksheet().addMergedRegion(new CellRangeAddress(rowStartIndex, rowEndIndex, colStartIndex, colEndIndex));
        };
    }
}
