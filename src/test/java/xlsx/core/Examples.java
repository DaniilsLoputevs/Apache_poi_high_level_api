package xlsx.core;

import lombok.val;
import models.User;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.Before;
import org.junit.Test;
import xlsx.utils.IOHelperTest;
import xlsx.utils.RandomTestDataGenerator;

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

/**
 * TODO : support parallel write for big amount of data.(more that 1000 elements in block)
 */
public class Examples {
    private static final String DIR_PATH_XLSX_TEST = "C:/Danik/DEVELOPMENT/TM2-dev-excel/xlsx-api-test";
    
    private final RandomTestDataGenerator random = new RandomTestDataGenerator();
    private final IOHelperTest ioHelperTest = new IOHelperTest();
    
    private Iterable<User> users;
    
    @Before
    public void init() {
        System.out.println("Start generate random data");
        users = random.genRandomUsers(50);
        System.out.println("Finish generate random data");
    }
    
    @Test
    public void easy() {
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
        
        ioHelperTest.toDiskFile(DIR_PATH_XLSX_TEST, bytes);
    }
    
    @Test
    public void hard() {
        val book = new ExcelBook();
        val headerStyle = book.makeStyle()
                .foregroundColor(ORANGE).fillPattern(SOLID_FOREGROUND)
                .verticalAlignment(VerticalAlignment.CENTER).horizontalAlignment(CENTER)
                .font(book.makeFont().fontName("Arial").bold(true).build())
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
        
        ioHelperTest.toDiskFile(DIR_PATH_XLSX_TEST, bytes);
    }
    
}
