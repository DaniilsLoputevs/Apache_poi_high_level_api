package utils;

import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;

import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class IOHelper {
    private static final DateTimeFormatter LOCAL_DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd__HH-mm");
    
    @SneakyThrows
    public void toDiskFile(String dirPath, byte[] fileBytes) {
        val fileName = dirPath + "/" + LocalDateTime.now().format(LOCAL_DATE_TIME_FORMATTER) + ".xlsx";
        @Cleanup val outputStream = new FileOutputStream(fileName);
        outputStream.write(fileBytes);
    }
}
