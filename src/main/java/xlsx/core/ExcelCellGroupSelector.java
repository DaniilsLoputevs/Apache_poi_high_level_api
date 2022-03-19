package xlsx.core;

import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.val;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Supplier;

/**
 * @author Daniils Loputevs
 */
@RequiredArgsConstructor
public class ExcelCellGroupSelector {
    @Getter
    private final ExcelCellGroupType type;
    /** key - groupName && val - group */
    private final Map<String, ExcelCellGroup> groups = new LinkedHashMap<>();
    private final String collectPattern;
    
    private int lastRowIndex;
    private int lastColIndex;
    ExcelDataBlock<?> innerDataBlock;
    
    public ExcelCellGroupSelector add(String groupName, Consumer<List<ExcelCell>> operation) {
        var group = groups.get(groupName);
        if (group == null) {
            group = new ExcelCellGroup(groupName);
//            group.innerDataBlock = this.innerDataBlock;
            group.addOperation(operation);
            groups.put(groupName, group);
        } else group.addOperation(operation);
        return this;
    }
    
    public ExcelCellGroupSelector add(String groupName, BiConsumer<ExcelDataBlock<?>, List<ExcelCell>> biOperation) {
        var group = groups.get(groupName);
        if (group == null) {
            group = new ExcelCellGroup(groupName);
            group.addOperation(biOperation);
            groups.put(groupName, group);
        } else group.addOperation(biOperation);
        return this;
    }
    
    
    public void collectCells() {
        int rowIndex = 0, colIndex = 0;
        for (val patternLine : collectPattern.split("\r\n")) {
            for (val cellIdentifier : patternLine.split(" ")) {
//                System.out.printf("collectCells :: cellIdentifier = [\"%s\"](%s:%s)%n", cellIdentifier, rowIndex, colIndex);
                
                var group = this.groups.get(cellIdentifier);
                if (group != null) group.addCell(new ExcelCell(rowIndex, colIndex));
                else {
                    group = new ExcelCellGroup(cellIdentifier, new ExcelCell(rowIndex, colIndex));
                    groups.put(cellIdentifier, group);
                }
                
                colIndex++;
            }
            colIndex = 0;
            rowIndex++;
        }
        groups.values().forEach(group -> group.innerDataBlock = this.innerDataBlock);
//        for (val entry : groups.entrySet()) {
//            System.out.println("collectCells :: " + entry.getKey() + " && " + entry.getValue().getPhantomCells().size());
//        }
    }
    
    /* package private */
    
    int terminateInnerCells(Sheet sheet, int rowOffset, Supplier<CellStyle> createCellStyle) {
        for (val excelCellGroup : groups.values()) {
            excelCellGroup.initInnerCells(sheet, rowOffset);
            lastRowIndex = Math.max(lastRowIndex, excelCellGroup.getLastRowIndex());
            lastColIndex = Math.max(lastColIndex, excelCellGroup.getLastColIndex());
            excelCellGroup.executeAllOperations();
            excelCellGroup.terminateAllCells(createCellStyle);
        }
        return lastRowIndex;
    }
    
}
