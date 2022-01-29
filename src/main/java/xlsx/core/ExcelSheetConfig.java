package xlsx.core;

import xlsx.utils.Pair;

import java.util.ArrayList;
import java.util.List;

public class ExcelSheetConfig {
    private final List<Pair<Integer, Double>> columnIndexAndWidth = new ArrayList<>();
    
    public ExcelSheetConfig add(Pair<Integer, Double> colWidth) {
        columnIndexAndWidth.add(colWidth);
        return this;
    }
    
}
