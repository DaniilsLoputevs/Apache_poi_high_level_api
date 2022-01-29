package xlsx.core;

import lombok.Getter;

import java.util.ArrayList;
import java.util.List;

@Getter
public class ExcelSheet {
    private final List<ExcelDataBlock<?>> dataBlocks = new ArrayList<>();
    private final List<ExcelSheetConfig> settings = new ArrayList<>();
    int maxColumnsCount;
    int totalRowsCount;
    
    public ExcelSheet add(ExcelDataBlock<?> block) {
        dataBlocks.add(block);
        return this;
    }
    
    public ExcelSheet add(ExcelSheetConfig config) {
        this.settings.add(config);
        return this;
    }
}
