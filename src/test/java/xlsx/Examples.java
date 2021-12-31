package xlsx;

import lombok.val;
import models.User;
import org.junit.Test;
import utils.IOHelper;
import utils.RandomTestDataGenerator;

import static java.awt.Color.YELLOW;
import static org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND;

/**
 * TODO : sort by merging
 * TODO : support CompletableFuture<Iterable<T>> in income data, not only Iterable<T>.
 * TODO : nextSheet() && nextBlock()
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
        val headerStyle = book.buildStyle()
                .foregroundColor(YELLOW).fillPattern(SOLID_FOREGROUND).build();
        val dateStyle = book.buildStyle("dd.MM.yy HH:mm").build();
        val idStyle = book.buildIdStyle();
        val amountStyle = book.buildCurrencyStyle();
        
        val bytes = book.addBlock(new ExcelBlock<>(users)
                .addDefaultHeader(headerStyle)
                .addColumn("ID", User::getId, idStyle)
                .addColumn("Name", User::getName)
                .addColumn("Role", User::getRole)
                .addColumn("Register Date", User::getRegisterDate, dateStyle)
                .addColumn("Active", User::isActive)
                .addColumn("Balance", User::getBalance, amountStyle)
        ).toBytes();
        
        System.out.println("Start write to disk");
        ioHelper.toDiskFile(DIR_PATH_XLSX_TEST, bytes);
        System.out.println("Finish write to disk");
    }
}
