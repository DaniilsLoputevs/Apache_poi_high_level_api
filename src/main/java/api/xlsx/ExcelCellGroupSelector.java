package api.xlsx;

import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.Setter;
import lombok.val;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.*;
import java.util.function.Consumer;

@RequiredArgsConstructor
public class ExcelCellGroupSelector<D> {
    @Getter
    private final ExcelCellGroupType type;
    private final Map<String, List<ExcelCell>> groups = new LinkedHashMap<>();
    private final Map<String, List<Consumer<List<ExcelCell>>>> groupsFuncs = new HashMap<>();
    private final String collectPattern;
    
    /** для внутренних нужд. */
    @Setter
    private ExcelDataBlock<D> excelDataBlockRef;
    private int lastRowIndex;
    private int lastColIndex;
    
    
    public ExcelCellGroupSelector<D> doForGroup(String groupName, Consumer<List<ExcelCell>> doForGroupFunc) {
        var funcList = groupsFuncs.get(groupName);
        if (funcList == null) {
            funcList = new ArrayList<Consumer<List<ExcelCell>>>();
            funcList.add(doForGroupFunc);
            groupsFuncs.put(groupName, funcList);
        } else {
            funcList.add(doForGroupFunc);
        }
        return this;
    }
    
    public ExcelCellGroupSelector<D> setValueAndHeaderForGroup(String groupName, String value, ExcelCellStyle headerStyle) {
        Consumer<List<ExcelCell>> doFunc = ((List<ExcelCell> cells) -> {
            if (cells.isEmpty()) return;
            for (val cell : cells) {
                cell.setValue(value);
                cell.setStyle(headerStyle);
            }
        });
        doForGroup(groupName, doFunc);
        return this;
    }
    
    public ExcelCellGroupSelector<D> mergeCellGroupAndSetValueAndStyle(String groupName, String value, ExcelCellStyle style) {
        val valueFinal = value;
        val styleFinal = style;
        Consumer<List<ExcelCell>> doFunc = ((List<ExcelCell> cells) -> {
            if (cells.isEmpty()) return;
            
            int rowStartIndex = Integer.MAX_VALUE, rowEndIndex = 0, colStartIndex = Integer.MAX_VALUE, colEndIndex = 0;
            for (val cell : cells) {
                cell.setValue(valueFinal);
                cell.setStyle(styleFinal);
                rowStartIndex = Math.min(rowStartIndex, cell.getRowIndex());
                rowEndIndex = Math.max(rowEndIndex, cell.getRowIndex());
                
                colStartIndex = Math.min(colStartIndex, cell.getColIndex());
                colEndIndex = Math.max(colEndIndex, cell.getColIndex());
            }
//            System.out.printf("merge group :: mg[%s](%s-%s && %s-%s)%n",groupName, rowStartIndex, rowEndIndex, colStartIndex, colEndIndex);
            excelDataBlockRef.getSheet().addMergedRegion(new CellRangeAddress(rowStartIndex, rowEndIndex, colStartIndex, colEndIndex));
        });
        doForGroup(groupName, doFunc);
        return this;
    }
    
    /* package private */
    
    public void collectCells() {
        int rowIndex = 0, colIndex = 0;
        for (val patternLine : collectPattern.split("\r\n")) {
            for (val cellIdentifier : patternLine.split(" ")) {
//                System.out.printf("collectCells :: cellIdentifier = [\"%s\"](%s:%s)%n", cellIdentifier, rowIndex, colIndex);
                
                var cellGroup = this.groups.get(cellIdentifier);
                if (cellGroup != null) cellGroup.add(new ExcelCell(rowIndex, colIndex));
                else {
                    // Важно что бы этот List был Изменяемым.
                    cellGroup = new ArrayList<>();
                    cellGroup.add(new ExcelCell(rowIndex, colIndex));
                    this.groups.put(cellIdentifier, cellGroup);
                }
                colIndex++;
            }
            colIndex = 0;
            rowIndex++;
        }
    }
    
    int initInnerCells(XSSFSheet sheet, int rowOffset) {
        for (val excelCellGroup : groups.values()) {
            for (val excelCell : excelCellGroup) {
                val actualRowIndex = excelCell.getRowIndex() + rowOffset;
                var row = sheet.getRow(actualRowIndex);
                if (row == null) row = sheet.createRow(actualRowIndex);
                
                excelCell.setInnerCell(row.getCell(excelCell.getColIndex()));
                // обновляет rowIndex с offset для block - до реального index
                excelCell.setRowIndex(actualRowIndex);
                
                lastRowIndex = Math.max(lastRowIndex, actualRowIndex);
                lastColIndex = Math.max(lastColIndex, excelCell.getColIndex());
            }
        }
        return lastRowIndex;
    }
    
    void executeGroupFuncs() {
        for (val excelCellGroup : groups.entrySet()) {
            val groupFunc = groupsFuncs.get(excelCellGroup.getKey());
            if (groupFunc == null) continue;
            
            val cells = excelCellGroup.getValue();
            for (val func : groupFunc) {
                func.accept(cells);
            }
            cells.forEach(ExcelCell::terminate);
        }
    }
    
}
