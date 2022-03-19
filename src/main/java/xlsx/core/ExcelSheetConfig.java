package xlsx.core;

import xlsx.utils.Pair;

import java.util.HashMap;
import java.util.Map;

public class ExcelSheetConfig {
    final Map<Integer, Integer> columnsIndexAndWidth = new HashMap<>();
    String sheetName;
    
    public ExcelSheetConfig add(Pair<Integer, Integer> colWidth) {
        columnsIndexAndWidth.put(colWidth.getFirst(), colWidth.getSecond());
        return this;
    }
    
    public ExcelSheetConfig set(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }
    
}
