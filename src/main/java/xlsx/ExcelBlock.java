package xlsx;

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
import java.util.function.Function;

import static utils.DateUtil.toCalendar;
import static xlsx.ExcelCellGroupType.HEADER;

@RequiredArgsConstructor
public class ExcelBlock<D> {
    @Getter
    private final List<ExcelColumn> columns = new ArrayList<>();
    
    private final Iterable<D> data;
    private final Map<ExcelCellGroupType, ExcelCellGroupSelector<?>> allGroups = new HashMap<>();
    @Getter
    @Setter
    private XSSFSheet sheet;
    private ExcelCellStyle defaultHeaderStyle;
    
    
    public ExcelBlock<D> addDefaultHeader(ExcelCellStyle headerStyle) {
        this.defaultHeaderStyle = headerStyle;
        return this;
    }
    
    public ExcelBlock<D> addColumnEmptyHeader(Function<D, Object> dataGetter, ExcelCellStyle dataStyle) {
        columns.add(new ExcelColumn("", null, dataGetter, (__) -> dataStyle));
        return this;
    }
    
    /** Добавляет колнку с defaultHeaderStyle */
    public ExcelBlock<D> addColumn(String headerValue, Function<D, Object> dataGetter) {
        columns.add(new ExcelColumn(headerValue, (defaultHeaderStyle == null) ? null : defaultHeaderStyle.terminate(), dataGetter, (__) -> ExcelCellStyle.EMPTY));
        return this;
    }
    
    /** Добавляет колнку с defaultHeaderStyle */
    public ExcelBlock<D> addColumn(String headerValue, Function<D, Object> dataGetter, ExcelCellStyle dataStyle) {
        columns.add(new ExcelColumn(headerValue, (defaultHeaderStyle == null) ? null : defaultHeaderStyle.terminate(), dataGetter, (__) -> dataStyle));
        return this;
    }
    
    /** Добавляет колнку с defaultHeaderStyle */
    public ExcelBlock<D> addColumn(String headerValue, Function<D, Object> dataGetter, Function<D, ExcelCellStyle> dataStyleFunc) {
        columns.add(new ExcelColumn(headerValue, (defaultHeaderStyle == null) ? null : defaultHeaderStyle.terminate(), dataGetter, dataStyleFunc));
        return this;
    }
    
    public ExcelBlock<D> addColumn(String headerValue, ExcelCellStyle headerStyle, Function<D, Object> dataGetter) {
        columns.add(new ExcelColumn(headerValue, (headerStyle == null) ? null : headerStyle.terminate(), dataGetter, (__) -> ExcelCellStyle.EMPTY));
        return this;
    }
    
    public ExcelBlock<D> addColumn(String headerValue, ExcelCellStyle headerStyle, Function<D, Object> dataGetter, ExcelCellStyle dataStyle) {
        columns.add(new ExcelColumn(headerValue, (headerStyle == null) ? null : headerStyle.terminate(), dataGetter, (__) -> dataStyle));
        return this;
    }
    
    public ExcelBlock<D> addColumn(String headerValue, ExcelCellStyle headerStyle, Function<D, Object> dataGetter, Function<D, ExcelCellStyle> dataStyleFunc) {
        columns.add(new ExcelColumn(headerValue, (headerStyle == null) ? null : headerStyle.terminate(), dataGetter, dataStyleFunc));
        return this;
    }
    
    public ExcelBlock<D> addCellGroupsSelector(ExcelCellGroupSelector<D> selector) {
        selector.setExcelBlockRef(this);
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
                createCellAndSetValue(currentRow, cellIndex++, column.dataGetter.apply(currentRowData), column.dataStyle.apply(currentRowData).terminate());
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
                createCellAndSetValue(headerRow, cellIndex++, col.headerValue, col.headerCS);
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
    
    
    @RequiredArgsConstructor
    private class ExcelColumn {
        private final String headerValue;
        private final CellStyle headerCS;
        private final Function<D, Object> dataGetter;
        private final Function<D, ExcelCellStyle> dataStyle;
    }
    
}
