package api.xlsx;

import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.Setter;
import lombok.val;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

import static api.utils.DateUtil.toCalendar;
import static api.tools.ExcelCellStyles.DEFAULT;
import static api.xlsx.ExcelCellGroupType.HEADER;

@RequiredArgsConstructor
public class ExcelDataBlock<D> {
    @Getter
    private final List<ExcelColumn<D>> columns = new ArrayList<>();
    
    private final Iterable<D> data;
    private final Map<ExcelCellGroupType, ExcelCellGroupSelector<?>> allGroups = new HashMap<>();
    @Getter
    @Setter
    private XSSFSheet sheet;
    @Setter
    private ExcelCellStyle defaultHeaderStyle;
    
    
    public ExcelDataBlock<D> add(ExcelColumn<D> column) {
        columns.add(column);
        if (column.getHeaderStyle() == DEFAULT) column.setHeaderStyle(defaultHeaderStyle);
        return this;
    }
    
    
    public ExcelDataBlock<D> addCellGroupsSelector(ExcelCellGroupSelector<D> selector) {
        selector.setExcelDataBlockRef(this);
        selector.collectCells();
        allGroups.put(selector.getType(), selector);
        return this;
    }
    
    public void writeToWorkBookSheet(XSSFSheet sheet) {
        int rowIndex = (sheet.getLastRowNum() == 0) ? 0 : sheet.getLastRowNum() + 2;
        
        rowIndex = setBlockHeader(sheet, rowIndex);
        
        for (val currentRowData : data) {
            val currentRow = sheet.createRow(rowIndex++);
            int cellIndex = 0;
            for (val column : columns) {
                createCellAndSetValue(currentRow, cellIndex++,
                        column.getDataGetter().apply(currentRowData),
                        column.getDataStyle().apply(currentRowData).terminate());
            }
        }
    }
    
    private int setBlockHeader(XSSFSheet sheet, int rowOffset) {
        if (allGroups.containsKey(HEADER)) {
            val headerGroup = allGroups.get(HEADER);
            rowOffset = headerGroup.initInnerCells(sheet, rowOffset);
            rowOffset++;
            headerGroup.executeGroupFuncs();
            
        } else {
            val headerRow = sheet.createRow(rowOffset++);
            int cellIndex = 0;
            for (val col : columns) {
                createCellAndSetValue(headerRow, cellIndex++, col.getHeaderValue(), col.getHeaderStyle().terminate());
            }
        }
        return rowOffset;
    }
    
    private void createCellAndSetValue(XSSFRow row, int cellIndex, Object cellValue, CellStyle cellStyle) {
        val cell = row.getCell(cellIndex);
        
        if (cellValue == null) cell.setCellValue("");
        else if (cellValue instanceof String) cell.setCellValue((String) cellValue);
        else if (cellValue instanceof Number) cell.setCellValue(((Number) cellValue).doubleValue());
        else if (cellValue instanceof Boolean) cell.setCellValue((Boolean) cellValue);
        else if (cellValue instanceof Enum) cell.setCellValue(((Enum) cellValue).name());
        
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
