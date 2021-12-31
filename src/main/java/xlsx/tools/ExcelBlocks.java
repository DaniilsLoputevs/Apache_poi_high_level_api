package xlsx.tools;

import lombok.val;
import xlsx.core.ExcelCellStyle;
import xlsx.core.ExcelDataBlock;

import java.util.concurrent.CompletableFuture;


/**
 * @author Daniils Loputevs
 */
public class ExcelBlocks {
    public static <T> ExcelDataBlock<T> block(Iterable<T> data) {
        return new ExcelDataBlock<>(CompletableFuture.completedFuture(data));
    }
    
    public static <T> ExcelDataBlock<T> block(Iterable<T> data, ExcelCellStyle defaultHeaderStyle) {
        val block = new ExcelDataBlock<>(CompletableFuture.completedFuture(data));
        block.setDefaultHeaderStyle(defaultHeaderStyle);
        return block;
    }
    
    public static <T> ExcelDataBlock<T> block(CompletableFuture<Iterable<T>> dataFuture) {
        return new ExcelDataBlock<>(dataFuture);
    }
    
    public static <T> ExcelDataBlock<T> block(CompletableFuture<Iterable<T>> dataFuture, ExcelCellStyle defaultHeaderStyle) {
        val block = new ExcelDataBlock<>(dataFuture);
        block.setDefaultHeaderStyle(defaultHeaderStyle);
        return block;
    }
    
}
