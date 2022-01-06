package xlsx.core;

import lombok.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.concurrent.CompletableFuture;

import static xlsx.core.ExcelCellGroupType.HEADER;
import static xlsx.tools.ExcelCellStyles.DEFAULT;
import static xlsx.utils.DateUtil.toCalendar;

/**
 * @author Daniils Loputevs
 */
@RequiredArgsConstructor
public class ExcelDataBlock<D> {
    @Getter
    private final List<ExcelColumn<D>> columns = new ArrayList<>();
    
    private final CompletableFuture<Iterable<D>> dataFuture;
    private final Map<ExcelCellGroupType, ExcelCellGroupSelector> allGroups = new HashMap<>();
    @Getter
    @Setter
    private Sheet sheet;
    @Setter
    private ExcelCellStyle defaultHeaderStyle;
    
    
    public ExcelDataBlock<D> add(ExcelColumn<D> column) {
        columns.add(column);
        if (column.getHeaderStyle() == DEFAULT) column.setHeaderStyle(defaultHeaderStyle);
        return this;
    }
    
    public ExcelDataBlock<D> add(ExcelCellGroupSelector selector) {
        selector.collectCells();
        allGroups.put(selector.getType(), selector);
        return this;
    }
    
    @SneakyThrows
    public void writeToWorkBookSheet(Sheet sheet) {
        System.out.println("cond = " + (sheet.getLastRowNum() == 0));
        System.out.println("sheet.getLastRowNum() = " + (sheet.getLastRowNum()));
        // if this dataBlock isn't first, we skip 1 empty line
        int rowIndex = (sheet.getLastRowNum() == -1) ? 0 : sheet.getLastRowNum() + 2;
    
        System.out.println("init rowIndex = " + rowIndex);
        rowIndex = setBlockHeader(sheet, rowIndex);
        
        for (val currentRowData : dataFuture.get()) {
            val currentRow = sheet.createRow(rowIndex++);
            int cellIndex = 0;
            for (val column : columns) {
                createCellAndSetValue(currentRow, cellIndex++,
                        column.getDataGetter().apply(currentRowData),
                        column.getDataStyle().apply(currentRowData).terminate());
            }
        }
    }

    private int setBlockHeader(Sheet sheet, int rowOffset) {
        if (allGroups.containsKey(HEADER)) {
            val headerGroup = allGroups.get(HEADER);
            rowOffset = headerGroup.initInnerCells(sheet, rowOffset);
            rowOffset++;
            
        } else {
            System.out.println("header rowOffset = " + rowOffset);
            val headerRow = sheet.createRow(rowOffset++);
            int cellIndex = 0;
            for (val col : columns) {
                createCellAndSetValue(headerRow, cellIndex++, col.getHeaderValue(), col.getHeaderStyle().terminate());
            }
        }
        return rowOffset;
    }
    
    private void createCellAndSetValue(Row row, int cellIndex, Object cellValue, CellStyle cellStyle) {
        val cell = row.getCell(cellIndex);
        
        if (cellValue == null) cell.setCellValue("");
        else if (cellValue instanceof String) cell.setCellValue((String) cellValue);
        else if (cellValue instanceof Number) cell.setCellValue(((Number) cellValue).doubleValue());
        else if (cellValue instanceof Boolean) cell.setCellValue((Boolean) cellValue);
        else if (cellValue instanceof Enum) cell.setCellValue(((Enum<?>) cellValue).name());
        
        else if (cellValue instanceof Calendar) cell.setCellValue((Calendar) cellValue);
        else if (cellValue instanceof Date) cell.setCellValue(toCalendar((Date) cellValue));
        else if (cellValue instanceof LocalDate) cell.setCellValue(toCalendar((LocalDate) cellValue));
        else if (cellValue instanceof LocalDateTime) cell.setCellValue(toCalendar((LocalDateTime) cellValue));
        else {
            System.out.println("WARM! cell value : try to set unsupported type: " + cellValue.getClass().getSimpleName());
            cell.setCellValue(cellValue.toString());
        }
        
        if (cellStyle != null) cell.setCellStyle(cellStyle);
    }
    
}
