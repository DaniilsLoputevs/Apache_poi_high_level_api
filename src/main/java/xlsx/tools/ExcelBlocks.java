package xlsx.tools;

import xlsx.core.ExcelCellStyle;
import xlsx.core.ExcelDataBlock;
import lombok.val;


/**
 * @author Daniils Loputevs
 */
public class ExcelBlocks {
    public static <T> ExcelDataBlock<T> block(Iterable<T> data, ExcelCellStyle defaultHeaderStyle) {
        val block = new ExcelDataBlock<>(data);
        block.setDefaultHeaderStyle(defaultHeaderStyle);
        return block;
    }
    
}
