package xlsx.core;

import lombok.Cleanup;
import lombok.Setter;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.StreamSupport;

import static xlsx.core.ExcelCellGroupType.HEADER;
import static xlsx.utils.DateUtil.toCalendar;

/**
 * Terminate whole Excel book to bytes.
 */
@Setter
public class ExcelBookWriter {
    private int cellCountToUseSXSSF = 2_000;
    
    
    public byte[] writeExcelBookToBytes(ExcelBook book) {
        return (book.isTerminated) ? toBytes(book) : toBytes(terminateExcelBook(book));
    }
    
    @SneakyThrows
    private byte[] toBytes(ExcelBook book) {
        @Cleanup val bos = new ByteArrayOutputStream();
        book.getWorkbook().write(bos);
        return bos.toByteArray();
    }
    
    public ExcelBook terminateExcelBook(ExcelBook book) {
        val useSXSSF = bookTotalCellCount(book) >= cellCountToUseSXSSF;
        book.workbook = useSXSSF ? new XSSFWorkbook() : new SXSSFWorkbook();
        
        for (val sheet : book.sheets) {
            val columnsMaxCharCount = new HashMap<Integer, Integer>();
            sheet.innerWorksheet = book.getWorkbook().createSheet(sheet.name);
            val innerSheet = sheet.innerWorksheet;
            
            for (val dataBlock : sheet.getDataBlocks()) {
                dataBlockWrite(dataBlock, innerSheet, columnsMaxCharCount);
            }
            
            // todo - set column width for sheet.
            //   XSSF & autosize ||  XSSF & hardcode width
            //  SXSSF & autosize || SXSSF & hardcode width
            for (int columnIndex = 0; columnIndex <= sheet.maxColumnsCount; columnIndex++) {
                if (!useSXSSF) innerSheet.autoSizeColumn(columnIndex);
                else innerSheet.setColumnWidth(columnIndex, columnsMaxCharCount.get(columnIndex));
                sheet.config.columnsIndexAndWidth.forEach(innerSheet::setColumnWidth);
            }
        }
        book.isTerminated = true;
        return book;
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
                sheet.totalRowsCount += utilsSizeOfIterable(dataBlock.getData());
            }
            maxColumnsCount = Math.max(sheet.maxColumnsCount, maxColumnsCount);
            totalRowsCount += sheet.totalRowsCount;
        }
        System.out.println("bookTotalCellCount#return = " + maxColumnsCount * totalRowsCount);
        return maxColumnsCount * totalRowsCount;
    }
    
    private <T> void dataBlockWrite(ExcelDataBlock<T> dataBlock, Sheet worksheet,
                                    Map<Integer, Integer> columnsMaxCharCount) {
        // if this dataBlock isn't first, we skip 1 empty line
        int rowIndex = (worksheet.getLastRowNum() == -1) ? 0 : worksheet.getLastRowNum() + 2;
        
        rowIndex = dataBlocWriteHeader(dataBlock, worksheet, rowIndex, columnsMaxCharCount);
        dataBlocWriteBody(dataBlock, worksheet, rowIndex, columnsMaxCharCount);
    }
    
    private <T> int dataBlocWriteHeader(ExcelDataBlock<T> dataBlock, Sheet worksheet, int rowOffset,
                                        Map<Integer, Integer> columnsMaxCharCount) {
        val headerGroup = dataBlock.allGroups.get(HEADER);
        if (headerGroup != null) {
            rowOffset = headerGroup.initInnerCells(worksheet, rowOffset);
            rowOffset++;
        } else {
            val headerRow = worksheet.createRow(rowOffset++);
            int cellIndex = 0;
            int columnIndex = 0;
            for (val column : dataBlock.columns) {
                val columnMaxCharCount = writeCell(headerRow, cellIndex++,
                        column.getHeaderValue(), column.getHeaderStyle().terminate());
                putMaxColumnCharCount(columnsMaxCharCount, columnIndex, columnMaxCharCount);
            }
        }
        return rowOffset;
    }
    
    private <T> void dataBlocWriteBody(ExcelDataBlock<T> dataBlock, Sheet worksheet, int rowIndex,
                                       Map<Integer, Integer> columnsMaxCharCount) {
        for (val currentRowData : dataBlock.getData()) {
            val currentRow = worksheet.createRow(rowIndex++);
            int cellIndex = 0;
            int columnIndex = 0;
            for (val column : dataBlock.columns) {
                val columnMaxCharCount = writeCell(currentRow, cellIndex++,
                        column.getDataGetter().apply(currentRowData),
                        column.getDataStyle().apply(currentRowData).terminate());
                putMaxColumnCharCount(columnsMaxCharCount, columnIndex, columnMaxCharCount);
            }
        }
    }
    
    private int writeCell(Row row, int cellIndex, Object cellValue, CellStyle cellStyle) {
        // todo - разобраться почему перестал работать .getCell()
//        val cell = row.getCell(cellIndex);
        val cell = row.createCell(cellIndex);
        int cellCharCount;
        System.out.println("row = " + row);
        System.out.println("cell = " + cell);
        
        if (cellValue == null) cellCharCount = setCellValueString(cell, "");
        else if (cellValue instanceof String) cellCharCount = setCellValueString(cell, (String) cellValue);
        else if (cellValue instanceof Number) cellCharCount = setCellValueNumber(cell, (Number) cellValue);
        else if (cellValue instanceof Boolean) cellCharCount = setCellValueBoolean(cell, (Boolean) cellValue);
        else if (cellValue instanceof Enum<?>) cellCharCount = setCellValueString(cell, ((Enum<?>) cellValue).name());
        
        else if (cellValue instanceof Calendar)
            cellCharCount = setCellValueCalendar(cell, (Calendar) cellValue, cellStyle);
        else if (cellValue instanceof Date)
            cellCharCount = setCellValueCalendar(cell, toCalendar((Date) cellValue), cellStyle);
        else if (cellValue instanceof LocalDate)
            cellCharCount = setCellValueCalendar(cell, toCalendar((LocalDate) cellValue), cellStyle);
        else if (cellValue instanceof LocalDateTime)
            cellCharCount = setCellValueCalendar(cell, toCalendar((LocalDateTime) cellValue), cellStyle);
        else {
            System.out.println("WARM! cell value : try to set unsupported type: " + cellValue.getClass().getSimpleName());
            val temp = cellValue.toString();
            cell.setCellValue(temp);
            cellCharCount = temp.length();
        }
        
        // todo - fix:
        //  java.lang.IllegalArgumentException:
        //  This Style does not belong to the supplied Workbook Styles Source.
        //  Are you trying to assign a style from one workbook to the cell of a different workbook?
        if (cellStyle != null) cell.setCellStyle(cellStyle);
        return cellCharCount;
    }
    
    private <C> int setCellValueString(Cell cell, String cellValue) {
        cell.setCellValue(cellValue);
        return cellValue.length();
    }
    
    private <C> int setCellValueNumber(Cell cell, Number cellValue) {
        val temp = cellValue.doubleValue();
        cell.setCellValue(temp);
        return cellValue.toString().length();
    }
    
    private <C> int setCellValueBoolean(Cell cell, Boolean cellValue) {
        cell.setCellValue(cellValue);
        return cellValue.toString().length();
    }
    
    private <C> int setCellValueCalendar(Cell cell, Calendar cellValue, CellStyle cellStyle) {
        cell.setCellValue(cellValue);
        return cellStyle == null
                ? cellValue.toString().length()
                : cellStyle.getDataFormatString().length();
    }
    
    private void putMaxColumnCharCount(Map<Integer, Integer> columnsMaxCharCount,
                                       int columnIndex, int columnMaxCharCount) {
        columnsMaxCharCount.merge(columnIndex, columnMaxCharCount, (a, b) -> Math.max(b, a));
    }
//    private Workbook initExcelBook(ExcelBook book) {
//        // todo - impl or remove
//    }
    
    /* local Utils */
    private static int utilsSizeOfIterable(Iterable<?> iterable) {
        return (iterable instanceof Collection)
                ? ((Collection<?>) iterable).size() // List, Set, Que, Deque
                : (int) StreamSupport.stream(iterable.spliterator(), false).count();
    }
    
}
