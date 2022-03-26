package xlsx.core;

import lombok.Cleanup;
import lombok.Getter;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.Workbook;

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
    
    
    void toFile(String filePath);
    
    File toFile(File file);
    
    byte[] toBytes();
    
    ExcelBook terminate();
    
    
    /* getters && setters */
    
    
    List<ExcelSheet> getSheets();
    
    Workbook getBook();
    
    Workbook setBook(Workbook workbook);
    
    boolean isTerminated();
    
    boolean isTerminated(boolean value);
    
}

/**
 * @author Daniils Loputevs
 */
@Getter
class ExcelBookImpl implements ExcelBook {
    private final ExcelBookWriter writer = new ExcelBookWriterImpl();
    private final List<ExcelSheet> sheets = new ArrayList<>();
    private final List<ExcelCellStyle> cellStyles = new ArrayList<>();
    private final List<ExcelFont> fonts = new ArrayList<>();
    boolean isTerminated = false;
    private Workbook workbook;
    
    @Override
    public ExcelBook add(ExcelSheet sheet) {
        sheet.name = "sheet " + sheets.size() + 1;
        sheets.add(sheet);
        return this;
    }
    
    @Override
    public ExcelCellStyle.ExcelCellStyleBuilder makeStyle() {
        return ExcelCellStyle.builder()
                .horizontalAlignment(RIGHT)
//                .dataFormatHelper(workbook.getCreationHelper().createDataFormat())
                ;
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
    
    
    /* getters && setters */
    
    
    @Override
    public Workbook getBook() {
        return workbook;
    }
    
    @Override
    public Workbook setBook(Workbook workbook) {
        this.workbook = workbook;
        return workbook;
    }
    
    @Override
    public boolean isTerminated() {
        return isTerminated;
    }
    
    @Override
    public boolean isTerminated(boolean value) {
        this.isTerminated = value;
        return value;
    }
}
