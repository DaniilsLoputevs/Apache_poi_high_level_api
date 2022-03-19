package xlsx;

import lombok.val;
import models.User;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.junit.Before;
import org.junit.Test;
import xlsx.core.ExcelBook;
import xlsx.utils.RandomTestDataGenerator;
import xlsx.utils.TimeMarker;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import static java.awt.Color.*;
import static org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER;
import static xlsx.core.ExcelBookWriter.BLANK_SPACE_OFFSET;
import static xlsx.core.ExcelCellGroupType.HEADER;
import static xlsx.tools.ExcelBlocks.block;
import static xlsx.tools.ExcelCellGroupSelectors.*;
import static xlsx.tools.ExcelCellStyles.buildCurrencyStyle;
import static xlsx.tools.ExcelCellStyles.buildIdStyle;
import static xlsx.tools.ExcelColumns.column;
import static xlsx.tools.ExcelColumns.columnNoHeader;
import static xlsx.tools.ExcelSheets.sheet;

/**
 * TODO : multi sheet write
 * TODO : Stream? - DENIED
 * TODO : Reactive? - DENIED
 * TODO : ? Global workbook, sheet Options|settings ?
 * TODO : ? Better AutoSize that default is. (helps: https://stackoverflow.com/questions/18983203/how-to-speed-up-autosizing-columns-in-apache-poi)
 * TODO : Sheet Setting - sheet name
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
 * TODO : 17 - make new Standard styles
 * TODO : 18 - renew examples tests
 * TODO : 19 -
 * TODO : NN - clean from @Deprecated API
 * TODO :
 *
 *
 * book < sheet < (options || dataBlock) < columns
 */
public class Examples {
    private static final DateTimeFormatter LOCAL_DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd__HH-mm");
    public static final String BASE_PATH = "C:/Danik/DEVELOPMENT/TM2-dev-excel/xlsx-api-test/";
    public static final String XLSX = ".xlsx";
    public static final String SIMPLE_OUTPUT_PATH = BASE_PATH + "1_simple/simple_" + currentTime() + XLSX;
    public static final String COMPLEX_OUTPUT_PATH = BASE_PATH + "2_complex/complex_" + currentTime() + XLSX;
    public static final String DEV_OUTPUT_PATH = BASE_PATH + "3_dev/dev_" + currentTime() + XLSX;
    
    private static String currentTime() {
        return LocalDateTime.now().format(LOCAL_DATE_TIME_FORMATTER);
    }
    
    
    private final RandomTestDataGenerator random = new RandomTestDataGenerator();
    
    private Iterable<User> users;
    
    @Before
    public void init() {
        System.out.println("Start generate random data");
        users = random.genRandomUsers(5_000);
//        users = random.genRandomUsers(1000);
        System.out.println("Finish generate random data");
        System.out.println(DEV_OUTPUT_PATH);
    }
    
    /**
     * XSSF
     * 50_000 * 6 = 300_000 cells
     * - 12 sec 14 ms
     * - 14 sec 46 ms
     * - 10 sec 66 ms
     * - 14 sec 44 ms
     * - 13 sec 97 ms
     * SXSSF
     * 50_000 * 6 = 300_000 cells
     * 1 - 6 sec 81 ms
     * 2 - 5 sec 48 ms
     * 3 - 5 sec 46 ms
     * 4 - 5 sec 37 ms
     * 5 - 5 sec 44 ms
     *
     *
     * XSSF
     * 5_000 * 6 = 30_000 cells
     * 1 - 4 sec 58 ms
     * 2 - 4 sec 88 ms
     * SXSSF
     * 5_000 * 6 = 30_000 cells
     * 1 - 5 sec 25 ms
     * 2 - 3 sec 50 ms
     * 3 - 3 sec 11 ms
     * 4 - 2 sec 784 ms
     * 5 - 3 sec 92 ms
     */
    @Test
    public void easy() {
        TimeMarker.addMark("Start xlsx");
        val book = new ExcelBook();
        val headerStyle = book.makeStyle().foregroundColor(YELLOW).fillPattern(SOLID_FOREGROUND).build();
        val dateStyle = book.makeStyle("dd.MM.yy HH:mm").build();
        
        book.add(block(users, headerStyle)
                .add(column("ID", User::getId, buildIdStyle(book)))
                .add(column("Name", User::getName))
                .add(column("Role", User::getRole))
                .add(column("Register Date", User::getRegisterDate, dateStyle))
                .add(column("Active", User::isActive))
                .add(column("Balance", User::getBalance, buildCurrencyStyle(book)))
        ).toFile(SIMPLE_OUTPUT_PATH);
        
        TimeMarker.addMark("Finish xlsx");
        TimeMarker.printMarks();
    }
    
    // TODO : BROKEN output file
    @Test
    public void complex() {
        val book = new ExcelBook();
        val headerStyle = book.makeStyle()
                .foregroundColor(ORANGE).fillPattern(SOLID_FOREGROUND).allSideAlignment(CENTER)
                .font(book.makeFont().fontName("Arial").bold(true).height(12).build())
                .borderAllSide(BorderStyle.THIN).build();
        val dateStyle = book.makeStyle("dd.MM.yy HH:mm").build();
        val idStyle = buildIdStyle(book);
        val amountStyle = buildCurrencyStyle(book);
        
        book.add(block(users)
                        .add(columnNoHeader(User::getId, idStyle))
                        .add(columnNoHeader(User::getName))
                        .add(columnNoHeader(User::getRole))
                        .add(columnNoHeader(User::getRegisterDate, dateStyle))
                        .add(columnNoHeader(User::isActive))
                        .add(columnNoHeader(User::getBalance, amountStyle))
                        .add(cellGroupSelector(HEADER, ""
                                        + "h h h h h h \r\n"
                                        + "1 2 3 4 5 6 \r\n")
//                        .add("h", mergeCellGroupAndSetValueAndStyle("All users report", headerStyle, book))
                                        .add("1", setValueAndHeaderForGroup("ID", headerStyle))
                                        .add("2", setValueAndHeaderForGroup("Name", headerStyle))
                                        .add("3", setValueAndHeaderForGroup("Role", headerStyle))
                                        .add("4", setValueAndHeaderForGroup("Register Date", headerStyle))
                                        .add("5", setValueAndHeaderForGroup("Active", headerStyle))
                                        .add("6", setValueAndHeaderForGroup("Balance", headerStyle))
                        )
        ).toFile(COMPLEX_OUTPUT_PATH);
        
        TimeMarker.addMark("Finish xlsx");
        TimeMarker.printMarks();
    }
    
    
    /**
     * 2359 - 2607 - 2481 - 2385 - 2337
     * 2556 - 2499 - 2760 - 2809 - 2367
     * 2577 - 2800 - 2565 - 3028 - 2261
     *
     * 2349 - 2900 - 2975 - 2532 - 2692
     * 2393 - 2900 - 2764 - 2547 - 2309
     * 2631 - 3122 - 2565 - 2509 - 2523
     *
     * 6411 - 5824 - 6003 - 6416 - 6057
     * 5576 - 5841 - 7368 - 6659 - 6173
     * 7116 - 7020 - 6254 - 5841 - 6051
     *
     * 3664 - 4054 - 3830 - 4456 - 4345
     * 3908 - 3644 - 3662 - 3891 - 4409
     * 4755 - 4217 - 4141 - 4226 - 3775
     *
     * Type    Rows    Cells    Mode            Time(sec)
     * SXSSF - 5_000 - 30_000 - Auto   - (2.23 - 3.00 (avg 2.56))
     * SXSSF - 5_000 - 30_000 - Const  - (2.31 - 3.12 (avg 2.65))
     * XSSF  - 5_000 - 30_000 - Auto   - (5.56 - 7.37 (avg 6.30))
     * XSSF  - 5_000 - 30_000 - Const  - (3.64 - 4.76 (avg 4.06))
     */
    @Test
    public void easyV1_3() {
        TimeMarker.addMark("Start xlsx");
        val book = new ExcelBook();
        val subHeaderStyle = book.makeStyle().foregroundColor(GREEN).fillPattern(SOLID_FOREGROUND).build();
        val headerStyle = book.makeStyle().foregroundColor(YELLOW).fillPattern(SOLID_FOREGROUND).allSideAlignment(CENTER).build();
        val dateStyle = book.makeStyle("dd.MM.yy HH:mm").build();
        
        book.add(sheet().add(block(users, subHeaderStyle)
                        .add(column("ID", User::getId, buildIdStyle(book)))
                        .add(column("Name", User::getName))
                        .add(column("Role", User::getRole))
                        .add(column("Register Date", User::getRegisterDate, dateStyle))
                        .add(column("Active", User::isActive))
//                .add(column("Balance", User::getBalance, buildCurrencyStyle(book)))
                        .add(column("Balance", User::getBalance))
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
    
}
