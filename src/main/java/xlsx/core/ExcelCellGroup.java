package xlsx.core;

import lombok.Data;
import lombok.val;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;

/**
 * @author Daniils Loputevs
 */
@Data
public class ExcelCellGroup {
    private final String groupName;
    /**
     * phantom cause at init stage this cells will have only coordinates, innerCell == null.
     * innerCell will be real created later.
     */
    private final List<ExcelCell> phantomCells = new ArrayList<>();
    /** operations that will be invoked then phantomCells will receive real value for innerCell. */
    private final List<Consumer<List<ExcelCell>>> operations = new ArrayList<>();
    
    
    private int lastRowIndex;
    private int lastColIndex;
    
    public ExcelCellGroup(String groupName) {
        this.groupName = groupName;
    }
    
    public ExcelCellGroup(String groupName, ExcelCell firstGroupCell) {
        this.groupName = groupName;
        addCell(firstGroupCell);
    }
    
    
    public void addCell(ExcelCell cell) {
        phantomCells.add(cell);
    }
    
    public void addOperation(Consumer<List<ExcelCell>> operation) {
        operations.add(operation);
    }
    
    void initInnerCells(Sheet sheet, int rowOffset) {
        if (phantomCells.isEmpty())
            throw new IllegalStateException(String.format("groupName=\"%s\" is empty! Check your code on GroupSelector", groupName));
        for (val phantomCell : phantomCells) {
            val actualRowIndex = phantomCell.getRowIndex() + rowOffset;
            var row = sheet.getRow(actualRowIndex);
            if (row == null) row = sheet.createRow(actualRowIndex);
            
            phantomCell.setInnerCell(row.getCell(phantomCell.getColIndex()));
            // обновляет rowIndex с offset для block - до реального index
            phantomCell.setRowIndex(actualRowIndex);
            
            lastRowIndex = Math.max(lastRowIndex, actualRowIndex);
            lastColIndex = Math.max(lastColIndex, phantomCell.getColIndex());
        }
    }
    
    void executeAllOperations() {
        for (val op : operations) {
            op.accept(phantomCells);
            phantomCells.forEach(ExcelCell::terminate);
        }
    }
}
