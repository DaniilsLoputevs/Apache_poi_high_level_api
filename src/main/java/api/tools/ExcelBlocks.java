package api.tools;

import api.xlsx.ExcelDataBlock;
import api.xlsx.ExcelCellStyle;
import lombok.val;


public class ExcelBlocks {
    public static<T> ExcelDataBlock<T> block(Iterable<T> data, ExcelCellStyle defaultHeaderStyle) {
        val block = new ExcelDataBlock<>(data);
        block.setDefaultHeaderStyle(defaultHeaderStyle);
        return block;
    }
    
}
