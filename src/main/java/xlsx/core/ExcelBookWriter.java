package xlsx.core;

import lombok.Cleanup;
import lombok.Setter;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.util.stream.StreamSupport;

/**
 * Terminate whole Excel book to bytes.
 */
@Setter
public class ExcelBookWriter {
    private int cellCountToUseSXSSF = 2_000;
    
    public byte[] writeExcelBookToBytes(ExcelBook book) {
        val useSXSSF = bookTotalCellCount(book) >= cellCountToUseSXSSF;
        book.setWorkbook(useSXSSF ? new XSSFWorkbook() : new SXSSFWorkbook());
        
        for (val sheet : book.getSheets()) {
            for (val dataBlock : sheet.getDataBlocks()) {
                // todo - считать макс кол-во символов.
                
            }
            // todo - set column width for sheet.
            //  XSSF & autosize ||  XSSF & hardcode width
            // SXSSF & autosize || SXSSF & hardcode width
        }
        
        return toBytes(book);
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
    
    private void writeExelSheetToBook(ExcelBook book, ExcelSheet sheet) {
    
    }
    
    @SneakyThrows
    private byte[] toBytes(ExcelBook book) {
        @Cleanup val bos = new ByteArrayOutputStream();
        book.getWorkbook().write(bos);
        return bos.toByteArray();
    }
    
    /* to local Utils */
    private static int utilsSizeOfIterable(Iterable<?> iterable) {
        // todo - сделать Более хитро, глянуть в CardParser logger
        // todo - вынести в отдельные Utils
        return (int) StreamSupport.stream(iterable.spliterator(), false).count();
    }
    
}
