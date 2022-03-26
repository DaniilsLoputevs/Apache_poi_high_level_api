package xlsx.core;

import lombok.Getter;
import org.apache.poi.ss.usermodel.Sheet;
import xlsx.utils.Pair;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Getter
public class ExcelSheet {
    final List<ExcelDataBlock<?>> dataBlocks = new ArrayList<>();
    final Map<Integer, Integer> columnsIndexAndWidth = new HashMap<>();
    String name;
    String sheetName;
    
    Sheet innerWorksheet;
    int maxColumnsCount;
    int totalRowsCount;
    
    public ExcelSheet add(ExcelDataBlock<?> block) {
        dataBlocks.add(block);
        return this;
    }
    
    public ExcelSheet add(Pair<Integer, Integer> colWidth) {
        columnsIndexAndWidth.put(colWidth.getFirst(), colWidth.getSecond());
        return this;
    }
}
