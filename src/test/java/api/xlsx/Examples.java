package api.xlsx;

import lombok.val;
import models.User;
import org.junit.Test;
import api.utils.IOHelper;
import api.utils.RandomTestDataGenerator;

import static api.tools.ExcelBlocks.block;
import static api.tools.ExcelCellStyles.buildCurrencyStyle;
import static api.tools.ExcelCellStyles.buildIdStyle;
import static api.tools.ExcelColumns.column;
import static java.awt.Color.YELLOW;
import static org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND;

/**
 * TODO : support CompletableFuture<Iterable<T>> in income data, not only Iterable<T>.
 * TODO : ? native cells ?
 * TODO : remove Generic from CellGroupSelector
 */
public class Examples {
    private static final String DIR_PATH_XLSX_TEST = "C:/Danik/DEVELOPMENT/TM2-dev-excel/xlsx-api-test";
    
    private final RandomTestDataGenerator random = new RandomTestDataGenerator();
    private final IOHelper ioHelper = new IOHelper();
    
    @Test
    public void simple() {
        System.out.println("Start generate random data");
        val users = random.genRandomUsers(50);
        System.out.println("Finish generate random data");
        
        val book = new ExcelBook();
        val headerStyle = book.makeStyle().foregroundColor(YELLOW).fillPattern(SOLID_FOREGROUND).build();
        val dateStyle = book.makeStyle("dd.MM.yy HH:mm").build();
        val idStyle = buildIdStyle(book);
        val amountStyle = buildCurrencyStyle(book);
        
        book.add(block(users, headerStyle)
                .add(column("ID", User::getId, idStyle))
                .add(column("Name",User::getName))
                .add(column("Role", User::getRole))
                .add(column("Register Date", User::getRegisterDate, dateStyle))
                .add(column("Active", User::isActive))
                .add(column("Balance", User::getBalance, amountStyle))
        ).toBytes();
        val bytes = book.toBytes();
        
        System.out.println("Start write to disk");
        ioHelper.toDiskFile(DIR_PATH_XLSX_TEST, bytes);
        System.out.println("Finish write to disk");
    }
}
