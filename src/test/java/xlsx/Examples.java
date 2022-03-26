package xlsx;

import lombok.val;
import models.User;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.junit.Before;
import org.junit.Test;
import xlsx.core.ExcelBook;
import xlsx.core.ExcelCellStyle;
import xlsx.utils.RandomTestDataGenerator;
import xlsx.utils.TimeMarker;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import static java.awt.Color.*;
import static org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER;
import static xlsx.core.ExcelBook.defaultBook;
import static xlsx.core.ExcelBookWriter.BLANK_SPACE_OFFSET;
import static xlsx.core.ExcelCellGroupType.HEADER;
import static xlsx.tools.ExcelBlocks.block;
import static xlsx.tools.ExcelCellGroupSelectors.*;
import static xlsx.tools.ExcelColumns.column;
import static xlsx.tools.ExcelColumns.columnNoHeader;
import static xlsx.tools.ExcelSheets.columnWidth;
import static xlsx.tools.ExcelSheets.sheet;

/**
 * TODO : Stream? - DENIED
 * TODO : Reactive? - DENIED
 * TODO : ? Better AutoSize that default is. (helps: https://stackoverflow.com/questions/18983203/how-to-speed-up-autosizing-columns-in-apache-poi)
 *
 * TODO : README - оставить ссылку на Examples(переместить из core в main package)
 * TODO : README - описать код в readme
 * TODO : README - итоговые xlsx файлы оставить в test Resource
 *
 * TODO : описать доку для DataBlock
 * TODO : описать доку для column
 * TODO : описать доку для pattern CellGroupSelector
 * TODO : потыкать SXSSF - для больше чем 5 000 записей
 * TODO :
 * TODO : FIX : if header not set, default use {@link xlsx.tools.ExcelCellStyles#EMPTY}
 * TODO :
 * TODO : 1 - multi-sheet ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 2 - separate BookWrite ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 3 - auto resolve SXSSF || XSSF ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 4 - autoSize for SXSSF + Optimization ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 5 - emptyHeaderColumn (rename)-> noHeaderColumn ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 6 - SheetConfig ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 7 - possible just terminate ExcelBook. (then use to other things) ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 8 - remove builder for ExcelCellStyle & ExcelFont + add to book them Make it. - DENIED OOOOOOOOOOOOOOOOOOOOOOO
 * TODO : 9 - terminate operation: toFile(String || File) ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 10 - (refactoring bug) создание workbook - потеряна настройка MISSING POLICY ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 11 - ExcelBookWriter make interface? Why not? ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 12 - ExcelSheetConfig.sheetName || ExcelSheet.name ??? ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 13 - CellStyle noStyle -> DEFAULT || EMPTY ???
 * TODO : 14 - dataBlock from CompletableFuture ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 15 - fix GroupSelector && merge region broken  ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 16 - fix GroupSelector && merge region not use styles at not first cell ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 17 - make new Standard styles - DENIED OOOOOOOOOOOOOOOOOOOOOOO000000000000000000000000000000000000000000000000
 * TODO : 18 - renew examples tests ???
 * TODO : 19 - make ExcelBook And ExcelBookWriter interface ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 20 - clean from @Deprecated API ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : NN -
 * TODO :
 */
public class Examples {
    public static final String BASE_PATH = "C:/Danik/DEVELOPMENT/TM2-dev-excel/xlsx-api-test/";
    public static final String XLSX = ".xlsx";
    private static final DateTimeFormatter LOCAL_DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd__HH-mm");
    public static final String SIMPLE_OUTPUT_PATH = BASE_PATH + "1_simple/simple_" + currentTime() + XLSX;
    public static final String COMPLEX_OUTPUT_PATH = BASE_PATH + "2_complex/complex_" + currentTime() + XLSX;
    public static final String DEV_OUTPUT_PATH = BASE_PATH + "3_dev/dev_" + currentTime() + XLSX;
    private final RandomTestDataGenerator random = new RandomTestDataGenerator();
    private Iterable<User> users;
    
    private static String currentTime() {
        return LocalDateTime.now().format(LOCAL_DATE_TIME_FORMATTER);
    }
    
    @Before
    public void init() {
        System.out.println("Start generate random data");
        users = random.genRandomUsers(5_000);
//        users = random.genRandomUsers(1000);
        System.out.println("Finish generate random data");
    }
    
    @Test
    public void easy() {
        TimeMarker.addMark("Start xlsx");
        val book = defaultBook();
        val headerStyle = book.makeStyle().foregroundColor(YELLOW).fillPattern(SOLID_FOREGROUND).build();
        val dateStyle = book.makeStyle("dd.MM.yy HH:mm").build();
        
        book.add(sheet().add(block(users, headerStyle)
                .add(column("ID", User::getId, buildIdStyle(book)))
                .add(column("Name", User::getName))
                .add(column("Role", User::getRole))
                .add(column("Register Date", User::getRegisterDate, dateStyle))
                .add(column("Active", User::isActive))
                .add(column("Balance", User::getBalance, buildCurrencyStyle(book)))
        )).toFile(SIMPLE_OUTPUT_PATH);
        
        TimeMarker.addMark("Finish xlsx");
        TimeMarker.printMarks();
    }
    
    @Test
    public void complex() {
        val book = defaultBook();
        val headerStyle = book.makeStyle()
                .foregroundColor(ORANGE).fillPattern(SOLID_FOREGROUND).allSideAlignment(CENTER)
                .font(book.makeFont().fontName("Arial").bold(true).height(12).build())
                .borderAllSide(BorderStyle.THIN).build();
        val dateStyle = book.makeStyle("dd.MM.yy HH:mm").build();
        val idStyle = buildIdStyle(book);
        val amountStyle = buildCurrencyStyle(book);
        
        book.add(sheet().add(block(users)
                .add(columnNoHeader(User::getId, idStyle))
                .add(columnNoHeader(User::getName))
                .add(columnNoHeader(User::getRole))
                .add(columnNoHeader(User::getRegisterDate, dateStyle))
                .add(columnNoHeader(User::isActive))
                .add(columnNoHeader(User::getBalance, amountStyle))
                .add(cellGroupSelector(HEADER, ""
                        + "h h h h h h \r\n"
                        + "1 2 3 4 5 6 \r\n")
                        .add("h", mergeCellGroupAndSetValueAndStyle("All users report", headerStyle, book))
                        .add("1", setValueAndHeaderForGroup("ID", headerStyle))
                        .add("2", setValueAndHeaderForGroup("Name", headerStyle))
                        .add("3", setValueAndHeaderForGroup("Role", headerStyle))
                        .add("4", setValueAndHeaderForGroup("Register Date", headerStyle))
                        .add("5", setValueAndHeaderForGroup("Active", headerStyle))
                        .add("6", setValueAndHeaderForGroup("Balance", headerStyle))
                ))
                .add(columnWidth(0, 2 + BLANK_SPACE_OFFSET))
        ).toFile(COMPLEX_OUTPUT_PATH);
        
        TimeMarker.addMark("Finish xlsx");
        TimeMarker.printMarks();
    }
    
    @Test
    public void easyV1_3() {
        TimeMarker.addMark("Start xlsx");
        val book = defaultBook();
        val subHeaderStyle = book.makeStyle().foregroundColor(GREEN).fillPattern(SOLID_FOREGROUND).build();
        val headerStyle = book.makeStyle().foregroundColor(YELLOW).fillPattern(SOLID_FOREGROUND).allSideAlignment(CENTER).build();
        val dateStyle = book.makeStyle("dd.MM.yy HH:mm").build();
        
        book.add(sheet().add(block(users, subHeaderStyle)
                        .add(column("ID", User::getId, buildIdStyle(book)))
                        .add(column("Name", User::getName))
                        .add(column("Role", User::getRole))
                        .add(column("Register Date", User::getRegisterDate, dateStyle))
                        .add(column("Active", User::isActive))
                        .add(column("Balance", User::getBalance, buildCurrencyStyle(book)))
                        .add(cellGroupSelector(HEADER, ""
                                + "h h h h h h \r\n"
                                + "1 2 3 4 5 6 \r\n")
                                .add("h", mergeCellGroupAndSetValueAndStyle("All users report", headerStyle, book))
                                .add("1", setValueAndHeaderForGroup("ID", subHeaderStyle))
                                .add("2", setValueAndHeaderForGroup("Name", subHeaderStyle))
                                .add("3", setValueAndHeaderForGroup("Role", subHeaderStyle))
                                .add("4", setValueAndHeaderForGroup("Register Date", subHeaderStyle))
                                .add("5", setValueAndHeaderForGroup("Active", subHeaderStyle))
                                .add("6", setValueAndHeaderForGroup("Balance", subHeaderStyle))
                        )
                )
//                .add(columnWidth(0, 2 + BLANK_SPACE_OFFSET))
//                .add(columnWidth(1, 4 + BLANK_SPACE_OFFSET))
//                .add(columnWidth(2, 5 + BLANK_SPACE_OFFSET))
//                .add(columnWidth(3, 14 + BLANK_SPACE_OFFSET))
//                .add(columnWidth(4, 6 + BLANK_SPACE_OFFSET))
//                .add(columnWidth(5, 22 + BLANK_SPACE_OFFSET))
        ).toFile(DEV_OUTPUT_PATH);
        
        TimeMarker.addMark("Finish xlsx");
        TimeMarker.printMarks();
    }
    
    private ExcelCellStyle buildIdStyle(ExcelBook excelBook) {
        return excelBook.makeStyle().format("0").build();
    }
    
    private ExcelCellStyle buildCurrencyStyle(ExcelBook excelBook) {
        return excelBook.makeStyle().format("0.00").build();
    }
    
}
