package api.xlsx;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

import java.util.function.Function;

// package private class
@AllArgsConstructor
@Getter
public class ExcelColumn<D> {
    private final String headerValue;
    @Setter
    private ExcelCellStyle headerStyle;
    private final Function<D, Object> dataGetter;
    private final Function<D, ExcelCellStyle> dataStyle;
}
