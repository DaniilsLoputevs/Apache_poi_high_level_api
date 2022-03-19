package xlsx.core;

import lombok.Setter;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.commons.collections4.IterableUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.function.Supplier;

import static xlsx.core.ExcelCellGroupType.HEADER;
import static xlsx.utils.DateUtil.toCalendar;

/** Terminate whole Excel book to bytes. */
public interface ExcelBookWriter {
    /** for more docs, see {@link Sheet#setColumnWidth} */
    int ABOUT_STANDARD_WIDTH_EXCEL_CHAR = 256;
    int BLANK_SPACE_OFFSET = 3;
    
    void writeExcelBookToOutput(ExcelBook book, OutputStream output);
    
    ExcelBook terminateExcelBook(ExcelBook book);
}

@Setter
class ExcelBookWriterImpl implements ExcelBookWriter{
    
    // todo - make normal value
    private int cellCountToUseSXSSF = 2_0000;
    
    @SneakyThrows
    public void writeExcelBookToOutput(ExcelBook book, OutputStream output) {
        if (book.isTerminated) book.workbook.write(output);
        else this.terminateExcelBook(book).workbook.write(output);
    }
    
    public ExcelBook terminateExcelBook(ExcelBook book) {
        val useSXSSF = bookTotalCellCount(book) >= cellCountToUseSXSSF;
        book.workbook = useSXSSF ? new SXSSFWorkbook() : new XSSFWorkbook();
        book.workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        
        for (val sheet : book.sheets) {
            val columnsMaxCharCount = new HashMap<Integer, Integer>();
            sheet.innerWorksheet = book.getWorkbook().createSheet(sheet.name);
            val innerSheet = sheet.innerWorksheet;
            
            sheet.getDataBlocks().forEach(it ->
                    dataBlockWrite(it, innerSheet, columnsMaxCharCount, book.workbook::createCellStyle));
            
            for (int columnIndex = 0; columnIndex < sheet.maxColumnsCount; columnIndex++) {
                /* set Const width || autosize */
                // todo - optimize for all cases : remove default autosize
                val constColumnWidth = sheet.columnsIndexAndWidth.get(columnIndex);
                if (constColumnWidth != null) innerSheet.setColumnWidth(columnIndex, constColumnWidth);
                else {
                    if (!useSXSSF) innerSheet.autoSizeColumn(columnIndex);
                    else autosizeColumnSXSSF(innerSheet, columnIndex, columnsMaxCharCount);
                }
            }
        }
        book.isTerminated = true;
        return book;
    }
    
    private void autosizeColumnSXSSF(Sheet innerSheet, int columnIndex, Map<Integer, Integer> columnsMaxCharCount) {
        val columnMaxCharCount = columnsMaxCharCount.get(columnIndex);
        innerSheet.setColumnWidth(columnIndex,
                columnMaxCharCount < 10
                        ? (BLANK_SPACE_OFFSET + columnMaxCharCount) * ABOUT_STANDARD_WIDTH_EXCEL_CHAR
                        : columnMaxCharCount * ABOUT_STANDARD_WIDTH_EXCEL_CHAR);
    }
    
    
    /**
     * It's need for understand how Big this file will be, to make decide: use SXSSF || XSSF.
     *
     * @param book -
     * @return total amount of real excel cells what will be used for write all data from all blocks.
     */
    private int bookTotalCellCount(ExcelBook book) {
        var maxColumnsCount = 0;
        var totalRowsCount = 0;
        for (val sheet : book.getSheets()) {
            for (val dataBlock : sheet.getDataBlocks()) {
                sheet.maxColumnsCount = Math.max(dataBlock.getColumns().size(), maxColumnsCount);
                sheet.totalRowsCount += IterableUtils.size(dataBlock.getData());
            }
            maxColumnsCount = Math.max(sheet.maxColumnsCount, maxColumnsCount);
            totalRowsCount += sheet.totalRowsCount;
        }
        System.out.println("bookTotalCellCount#return = " + maxColumnsCount * totalRowsCount);
        return maxColumnsCount * totalRowsCount;
    }
    
    private <T> void dataBlockWrite(ExcelDataBlock<T> dataBlock, Sheet worksheet,
                                    Map<Integer, Integer> columnsMaxCharCount,
                                    Supplier<CellStyle> createCellStyle) {
        dataBlock.setSheet(worksheet);
        /* if this dataBlock isn't first, we skip 1 empty line */
        int rowIndex = (worksheet.getLastRowNum() == -1) ? 0 : worksheet.getLastRowNum() + 2;
        
        rowIndex = dataBlockHeaderWrite(dataBlock, worksheet, rowIndex, columnsMaxCharCount, createCellStyle);
        dataBlockBodyWrite(dataBlock, worksheet, rowIndex, columnsMaxCharCount, createCellStyle);
    }
    
    private <T> int dataBlockHeaderWrite(ExcelDataBlock<T> dataBlock, Sheet worksheet,
                                         int rowOffset, Map<Integer, Integer> columnsMaxCharCount,
                                         Supplier<CellStyle> createCellStyle) {
        val headerGroup = dataBlock.allGroups.get(HEADER);
        if (headerGroup != null) {
            /* terminate && write value ti cells */
            rowOffset = headerGroup.terminateInnerCells(worksheet, rowOffset, createCellStyle);
            rowOffset++;
        } else {
            val headerRow = worksheet.createRow(rowOffset++);
            int cellIndex = 0;
            int columnIndex = 0;
            for (val column : dataBlock.columns) {
                val headerCellStyle = column.getHeaderStyle();
                val columnMaxCharCount = cellWrite(headerRow, cellIndex++,
                        column.getHeaderValue(), headerCellStyle, createCellStyle);
                putMaxColumnCharCount(columnsMaxCharCount, columnIndex++, columnMaxCharCount);
            }
        }
        return rowOffset;
    }
    
    private <T> void dataBlockBodyWrite(ExcelDataBlock<T> dataBlock, Sheet worksheet,
                                        int rowIndex, Map<Integer, Integer> columnsMaxCharCount,
                                        Supplier<CellStyle> createCellStyle) {
        for (val currentRowData : dataBlock.getData()) {
            val currentRow = worksheet.createRow(rowIndex++);
            int cellIndex = 0;
            int columnIndex = 0;
            for (val column : dataBlock.columns) {
                val dataCellStyle = column.getDataStyle().apply(currentRowData);
                val columnMaxCharCount = cellWrite(currentRow, cellIndex++,
                        column.getDataGetter().apply(currentRowData),
                        dataCellStyle, createCellStyle);
                putMaxColumnCharCount(columnsMaxCharCount, columnIndex++, columnMaxCharCount);
            }
        }
    }
    
    private int cellWrite(Row row, int cellIndex, Object cellValue, ExcelCellStyle excelCellStyle,
                          Supplier<CellStyle> createCellStyle) {
        val cellStyle = excelCellStyle.isTerminated
                ? excelCellStyle.cellStyleInner
                : excelCellStyle.terminate(createCellStyle.get());
        val cell = row.getCell(cellIndex);
        int cellChars;
        val displayFormat = excelCellStyle.getFormat();
        
        if (cellValue == null) cellChars = setCellValue(cell, "", displayFormat);
        else if (cellValue instanceof String) cellChars = setCellValue(cell, (String) cellValue, displayFormat);
        else if (cellValue instanceof Number) cellChars = setCellValue(cell, (Number) cellValue, displayFormat);
        else if (cellValue instanceof Boolean) cellChars = setCellValue(cell, (Boolean) cellValue);
        else if (cellValue instanceof Enum) cellChars = setCellValue(cell, ((Enum<?>) cellValue).name(), displayFormat);
        
        else if (cellValue instanceof Calendar)
            cellChars = setCellValue(cell, (Calendar) cellValue, displayFormat);
        else if (cellValue instanceof Date)
            cellChars = setCellValue(cell, toCalendar((Date) cellValue), displayFormat);
        else if (cellValue instanceof LocalDate)
            cellChars = setCellValue(cell, toCalendar((LocalDate) cellValue), displayFormat);
        else if (cellValue instanceof LocalDateTime)
            cellChars = setCellValue(cell, toCalendar((LocalDateTime) cellValue), displayFormat);
        else {
            System.out.println("WARM! cell value : try to set unsupported type: " + cellValue.getClass().getSimpleName());
            val temp = cellValue.toString();
            cell.setCellValue(temp);
            cellChars = temp.length();
        }
        
        if (cellStyle != null) cell.setCellStyle(cellStyle);
        return cellChars;
    }
    
    private <C> int setCellValue(Cell cell, String cellValue, String displayFormat) {
        cell.setCellValue(cellValue);
        return (displayFormat != null) ? displayFormat.length() : cellValue.length();
    }
    
    private <C> int setCellValue(Cell cell, Number cellValue, String displayFormat) {
        cell.setCellValue(cellValue.doubleValue());
        return (displayFormat != null) ? displayFormat.length() : cellValue.toString().length();
    }
    
    private <C> int setCellValue(Cell cell, Boolean cellValue) {
        cell.setCellValue(cellValue);
        return cellValue.toString().length();
    }
    
    private <C> int setCellValue(Cell cell, Calendar cellValue, String displayFormat) {
        cell.setCellValue(cellValue);
        return (displayFormat != null) ? displayFormat.length() : cellValue.toString().length();
    }
    
    private void putMaxColumnCharCount(Map<Integer, Integer> columnsMaxCharCount,
                                       int columnIndex, int columnMaxCharCount) {
        columnsMaxCharCount.merge(columnIndex, columnMaxCharCount, (a, b) -> Math.max(b, a));
    }
    
}
