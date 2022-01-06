package xlsx.core;

import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Cell;

/**
 * @author Daniils Loputevs
 */
@Setter
@Getter
@RequiredArgsConstructor
public class ExcelCell {
    private int rowIndex;
    private int colIndex;
    /**
     * На данный момент предполагается что сюда будут попадать только Sting,
     * т.к. это юзается только для block.header.
     * Когда будет потребность в других типах будет поддержка других типов.
     */
    private Object value;
    private ExcelCellStyle style;
    
    @Getter
    private Cell innerCell;
    
    public ExcelCell(int rowIndex, int colIndex) {
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
    }
    
    public Cell terminate() {
        if (innerCell == null) throw new IllegalStateException("innerCell is null");
        
        if (value != null) innerCell.setCellValue((String) value);
        if (style != null) innerCell.setCellStyle(style.terminate());
        return innerCell;
    }
    
    @Override
    public String toString() {
        return String.format("cell[\"%s\"](%s:%s)", value.toString(), rowIndex, colIndex);
    }
    
}
