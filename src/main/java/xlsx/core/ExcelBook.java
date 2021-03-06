package xlsx.core;

import lombok.Cleanup;
import lombok.Getter;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import xlsx.utils.Pair;

import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT;

/**
 * @author Daniils Loputevs
 */
public class ExcelBook {
    // TODO: work with SXXSFWorkbook
    private final Workbook workbook = new XSSFWorkbook();
    private final List<ExcelDataBlock<?>> blocks = new ArrayList<>();
    @Getter
    private final Sheet firstWorksheet = workbook.createSheet("sheet 1");
    
    private List<Pair<Integer, Integer>> globalColIndexes;
    
    public ExcelBook() {
        workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
    }
    
    
    public ExcelBook add(ExcelDataBlock<?> block) {
        block.setSheet(firstWorksheet);
        blocks.add(block);
        return this;
    }
    
    
    public ExcelBook globalSetColumnWidth(int colIndex, int width) {
        if (globalColIndexes == null) globalColIndexes = new ArrayList<>();
        globalColIndexes.add(new Pair<>(colIndex, width));
        return this;
    }
    
    public ExcelCellStyle.ExcelCellStyleBuilder makeStyle() {
        return ExcelCellStyle.builder()
                .cellStyleInner(workbook.createCellStyle())
                .horizontalAlignment(LEFT)
                .dataFormatHelper(workbook.getCreationHelper().createDataFormat());
    }
    
    /** @param format - {@link ExcelCellStyle} */
    public ExcelCellStyle.ExcelCellStyleBuilder makeStyle(String format) {
        return makeStyle().format(format);
    }
    
    
    public ExcelFont.ExcelFontBuilder makeFont() {
        return ExcelFont.builder().innerFont(workbook.createFont());
    }
    
    @SneakyThrows
    public byte[] toBytes() {
        blocks.forEach(block -> block.writeToWorkBookSheet(firstWorksheet));
        
        autoSizeAllColumns(firstWorksheet);
        
        if (globalColIndexes != null)
            globalColIndexes.forEach(pair -> firstWorksheet.setColumnWidth(pair.getFirst(), pair.getSecond()));
        
        @Cleanup val bos = new ByteArrayOutputStream();
        workbook.write(bos);
        
        return bos.toByteArray();
    }
    
    private void autoSizeAllColumns(Sheet sheet) {
        int lastColumnIndex = 0;
        for (val block : blocks) {
            lastColumnIndex = Math.max(lastColumnIndex, block.getColumns().size());
        }
        for (int columnIndex = 0; columnIndex < lastColumnIndex; columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
        }
    }
    
}
