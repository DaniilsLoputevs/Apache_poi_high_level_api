package xlsx;

import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utils.Pair;

import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT;

public class ExcelBook {
    private final XSSFWorkbook workbook = new XSSFWorkbook();
    private final List<ExcelBlock<?>> blocks = new ArrayList<>();
    private final XSSFSheet firstWorksheet = workbook.createSheet("sheet 1");
    
    private List<Pair<Integer, Integer>> globalColIndexes;
    
    public ExcelBook() {
        workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
    }
    
    public ExcelBook addBlock(ExcelBlock<?> block) {
        block.setSheet(firstWorksheet);
        blocks.add(block);
        return this;
    }
    
    public ExcelBook globalSetColumnWidth(int colIndex, int width) {
        if (globalColIndexes == null) globalColIndexes = new ArrayList<>();
        globalColIndexes.add(new Pair(colIndex, width));
        return this;
    }
    
    public ExcelCellStyle.ExcelCellStyleBuilder buildStyle() {
        return ExcelCellStyle.builder()
                .cellStyleInner(workbook.createCellStyle())
                .horizontalAlignment(LEFT)
                .dataFormatHelper(workbook.getCreationHelper().createDataFormat());
    }
    
    /** @param format - {@link ExcelCellStyle} */
    public ExcelCellStyle.ExcelCellStyleBuilder buildStyle(String format) {
        return buildStyle().format(format);
    }
    
    public ExcelCellStyle buildIdStyle() {
        return buildStyle().format("0").build();
    }
    
    public ExcelCellStyle buildCurrencyStyle() {
        return buildStyle().format("0.00").build();
    }
    
    public ExcelFont.ExcelFontBuilder buildFont() {
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
    
    private void autoSizeAllColumns(XSSFSheet sheet) {
        int lastColumnIndex = 0;
        for (val block : blocks) {
            lastColumnIndex = Math.max(lastColumnIndex, block.getColumns().size());
        }
        for (int columnIndex = 0; columnIndex < lastColumnIndex; columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
        }
    }
    
}
