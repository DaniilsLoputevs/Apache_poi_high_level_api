package xlsx;

import lombok.val;
import models.User;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.junit.Before;
import org.junit.Test;
import xlsx.core.ExcelBook;
import xlsx.utils.IOHelperTest;
import xlsx.utils.RandomTestDataGenerator;
import xlsx.utils.TimeMarker;

import static java.awt.Color.ORANGE;
import static java.awt.Color.YELLOW;
import static org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER;
import static xlsx.core.ExcelCellGroupType.HEADER;
import static xlsx.tools.ExcelBlocks.block;
import static xlsx.tools.ExcelCellGroupSelectors.*;
import static xlsx.tools.ExcelCellStyles.buildCurrencyStyle;
import static xlsx.tools.ExcelCellStyles.buildIdStyle;
import static xlsx.tools.ExcelColumns.column;
import static xlsx.tools.ExcelColumns.columnEmptyHeader;
import static xlsx.tools.ExcelSheetConfigs.columnWidth;
import static xlsx.tools.ExcelSheetConfigs.config;
import static xlsx.tools.ExcelSheets.sheet;

/**
 * TODO : multi sheet write
 * TODO : Stream?
 * TODO : Reactive?
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
 * TODO : потыкать SXSSF - для больше чем 7 500 записей
 * TODO :
 * TODO : FIX : if header not set, default use {@link xlsx.tools.ExcelCellStyles#EMPTY}
 * TODO :
 * TODO : 1 - multi-sheet ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 2 - separate BookWrite ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 3 - auto resolve SXSSF || XSSF ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 4 - autoSize for SXSSF
 * TODO : 5 - emptyHeaderColumn -> noHeaderColumn
 * TODO : 6 - SheetConfig ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 * TODO : 7 - possible just terminate ExcelBook. (then use to other things) ^^^^^^^^
 * TODO :
 *
 *
 * book < sheet < (options || dataBlock) < columns
 */
public class Examples {
    private static final String DIR_PATH_XLSX_TEST = "C:/Danik/DEVELOPMENT/TM2-dev-excel/xlsx-api-test";
    
    private final RandomTestDataGenerator random = new RandomTestDataGenerator();
    private final IOHelperTest ioHelperTest = new IOHelperTest();
    
    private Iterable<User> users;
    
    @Before
    public void init() {
        System.out.println("Start generate random data");
        users = random.genRandomUsers(5_000);
//        users = random.genRandomUsers(1000);
        System.out.println("Finish generate random data");
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
        
        val bytes = book.add(block(users, headerStyle)
                .add(column("ID", User::getId, buildIdStyle(book)))
                .add(column("Name", User::getName))
                .add(column("Role", User::getRole))
                .add(column("Register Date", User::getRegisterDate, dateStyle))
                .add(column("Active", User::isActive))
                .add(column("Balance", User::getBalance, buildCurrencyStyle(book)))
        ).toBytes();
        
        TimeMarker.addMark("Finish xlsx");
        TimeMarker.printMarks();
        
        ioHelperTest.toDiskFile(DIR_PATH_XLSX_TEST, bytes);
    }
    
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
        
        val bytes = book.add(block(users)
                .add(columnEmptyHeader(User::getId, idStyle))
                .add(columnEmptyHeader(User::getName))
                .add(columnEmptyHeader(User::getRole))
                .add(columnEmptyHeader(User::getRegisterDate, dateStyle))
                .add(columnEmptyHeader(User::isActive))
                .add(columnEmptyHeader(User::getBalance, amountStyle))
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
                )
        ).toBytes();
        
        TimeMarker.addMark("Finish xlsx");
        TimeMarker.printMarks();
        
        ioHelperTest.toDiskFile(DIR_PATH_XLSX_TEST, bytes);
    }
    
    
    @Test
    public void easyV1_3() {
        TimeMarker.addMark("Start xlsx");
        val book = new ExcelBook();
        val headerStyle = book.makeStyle().foregroundColor(YELLOW).fillPattern(SOLID_FOREGROUND).build();
        val dateStyle = book.makeStyle("dd.MM.yy HH:mm").build();
        
        val bytes = book.add(sheet().add(block(users, headerStyle)
                .add(column("ID", User::getId, buildIdStyle(book)))
                .add(column("Name", User::getName))
                .add(column("Role", User::getRole))
                .add(column("Register Date", User::getRegisterDate, dateStyle))
                .add(column("Active", User::isActive))
                .add(column("Balance", User::getBalance, buildCurrencyStyle(book)))
        ).set(config()
                .add(columnWidth(0, 12))
        )).toBytes();
        
        TimeMarker.addMark("Finish xlsx");
        TimeMarker.printMarks();
        
        ioHelperTest.toDiskFile(DIR_PATH_XLSX_TEST, bytes);
    }
    
}
