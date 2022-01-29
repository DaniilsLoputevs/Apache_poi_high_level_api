package xlsx.core;

import lombok.Getter;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;

@Getter
public class ExcelSheet {
    String name;
    final List<ExcelDataBlock<?>> dataBlocks = new ArrayList<>();
    ExcelSheetConfig config;
    Sheet innerWorksheet;
    int maxColumnsCount;
    int totalRowsCount;
    
    public ExcelSheet add(ExcelDataBlock<?> block) {
        dataBlocks.add(block);
        return this;
    }
    
    public ExcelSheet set(ExcelSheetConfig config) {
        this.config = config;
        return this;
    }
}
