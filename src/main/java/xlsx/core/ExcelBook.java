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
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT;


public interface ExcelBook {
    
    static ExcelBook defaultBook() {
        return new ExcelBookImpl();
    }
    
    ExcelBook add(ExcelSheet sheet);
    
    ExcelCellStyle.ExcelCellStyleBuilder makeStyle();
    
    /** @param format - {@link ExcelCellStyle} */
    ExcelCellStyle.ExcelCellStyleBuilder makeStyle(String format);
    
    ExcelFont.ExcelFontBuilder makeFont();
    
    
    /* Terminate operations */
    
    
    void toFile(String filePath);
    
    File toFile(File file);
    
    byte[] toBytes();
    
    ExcelBook terminate();
    
    List<ExcelSheet> getSheets();
    boolean terminated();
    boolean isTerminated(boolean value);
    
}

/**
 * @author Daniils Loputevs
 */
@Getter
class ExcelBookImpl implements ExcelBook {
    
    Workbook workbook;
    @Deprecated
    private final List<ExcelDataBlock<?>> blocks = new ArrayList<>();
    @Getter
    @Deprecated
    private Sheet firstWorksheet;
    @Deprecated
    private List<Pair<Integer, Integer>> globalColIndexes;
    
    final ExcelBookWriter writer = new ExcelBookWriterImpl();
    final List<ExcelSheet> sheets = new ArrayList<>();
    final List<ExcelCellStyle> cellStyles = new ArrayList<>();
    final List<ExcelFont> fonts = new ArrayList<>();
    boolean isTerminated = false;
    
    @Deprecated
    public ExcelBookImpl() {
        this.workbook = new XSSFWorkbook();
        init();
    }
    
    @Deprecated
    // TODO : быстрое решение, хотелось бы, сделать по лучше, чем такой полу-костыль.
    public ExcelBookImpl(Workbook workbook) {
        this.workbook = workbook;
        init();
    }
    
    @Deprecated
    private void init() {
        workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        this.firstWorksheet = workbook.createSheet("sheet 1");
    }
    
    
    /**
     * TODO : remove in release!!! Use add(ExcelSheet sheet)
     */
    @Deprecated
    public ExcelBook add(ExcelDataBlock<?> block) {
        block.setSheet(firstWorksheet);
        blocks.add(block);
        return this;
    }
    
    @Override
    public ExcelBook add(ExcelSheet sheet) {
        sheet.name = "sheet " + sheets.size() + 1;
        sheets.add(sheet);
        return this;
    }
    
    @Deprecated
    // ?O_O need ot not?
    /** This method is SPECIAL made package private. */
    void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }
    
    
    @Deprecated
    public ExcelBook globalSetColumnWidth(int colIndex, int width) {
        if (globalColIndexes == null) globalColIndexes = new ArrayList<>();
        globalColIndexes.add(new Pair<>(colIndex, width));
        return this;
    }
    
    @Override
    public ExcelCellStyle.ExcelCellStyleBuilder makeStyle() {
        val rsl = ExcelCellStyle.builder()
//                .cellStyleInner(workbook.createCellStyle())
                .horizontalAlignment(RIGHT)
                .dataFormatHelper(workbook.getCreationHelper().createDataFormat());
//        cellStyles.add(rsl);
        return rsl;
    }
    
    /** @param format - {@link ExcelCellStyle} */
    @Override
    public ExcelCellStyle.ExcelCellStyleBuilder makeStyle(String format) {
        return makeStyle().format(format);
    }
    
    
    @Override
    public ExcelFont.ExcelFontBuilder makeFont() {
        return ExcelFont.builder().innerFont(workbook.createFont());
    }
    
    
    /* Terminate operations */
    
    
    @SneakyThrows
    @Override
    public void toFile(String filePath) {
        @Cleanup val output = new FileOutputStream(filePath);
        writer.writeExcelBookToOutput(this, output);
    }
    
    @SneakyThrows
    @Override
    public File toFile(File file) {
        @Cleanup val output = new FileOutputStream(file);
        writer.writeExcelBookToOutput(this, output);
        return file;
    }
    
    @SneakyThrows
    @Override
    public byte[] toBytes() {
        @Cleanup val output = new ByteArrayOutputStream();
        writer.writeExcelBookToOutput(this, output);
        return output.toByteArray();
    }
    
    @Override
    public ExcelBook terminate() {
        return writer.terminateExcelBook(this);
    }
    
}
