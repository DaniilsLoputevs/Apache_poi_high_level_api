package xlsx.tools;

import xlsx.core.ExcelCellStyle;
import xlsx.core.ExcelColumn;

import java.util.function.Function;

import static xlsx.tools.ExcelCellStyles.DEFAULT;
import static xlsx.tools.ExcelCellStyles.EMPTY;

/**
 * @author Daniils Loputevs
 */
public final class ExcelColumns {
    
    /* empty header */
    
    public static <D> ExcelColumn<D> columnEmptyHeader(Function<D, Object> dataGetter, ExcelCellStyle dataStyle) {
        return new ExcelColumn<>("", EMPTY, dataGetter, (__) -> dataStyle);
    }
    
    public static <D> ExcelColumn<D> columnEmptyHeader(Function<D, Object> dataGetter) {
        return new ExcelColumn<>("", EMPTY, dataGetter, (__) -> EMPTY);
    }
    
    public static <D> ExcelColumn<D> columnEmptyHeader(Function<D, Object> dataGetter, Function<D, ExcelCellStyle> dataStyleFunc) {
        return new ExcelColumn<>("", EMPTY, dataGetter, dataStyleFunc);
    }
    
    /* default header style columns */
    
    public static <D> ExcelColumn<D> column(String headerValue, Function<D, Object> dataGetter) {
        return new ExcelColumn<>(headerValue, DEFAULT, dataGetter, (__) -> EMPTY);
    }
    
    public static <D> ExcelColumn<D> column(String headerValue, Function<D, Object> dataGetter, ExcelCellStyle dataStyle) {
        return new ExcelColumn<>(headerValue, DEFAULT, dataGetter, (__) -> dataStyle);
    }
    
    public static <D> ExcelColumn<D> column(String headerValue, Function<D, Object> dataGetter, Function<D, ExcelCellStyle> dataStyleFunc) {
        return new ExcelColumn<>(headerValue, DEFAULT, dataGetter, dataStyleFunc);
    }
    
    /* regular columns */
    
    public static <D> ExcelColumn<D> column(String headerValue, ExcelCellStyle headerStyle, Function<D, Object> dataGetter) {
        return new ExcelColumn<>(headerValue, (headerStyle == null) ? EMPTY : headerStyle, dataGetter, (__) -> EMPTY);
    }
    
    public static <D> ExcelColumn<D> column(String headerValue, ExcelCellStyle headerStyle, Function<D, Object> dataGetter, ExcelCellStyle dataStyle) {
        return new ExcelColumn<>(headerValue, (headerStyle == null) ? EMPTY : headerStyle, dataGetter, (__) -> dataStyle);
    }
    
    public static <D> ExcelColumn<D> column(String headerValue, ExcelCellStyle headerStyle, Function<D, Object> dataGetter, Function<D, ExcelCellStyle> dataStyleFunc) {
        return new ExcelColumn<>(headerValue, (headerStyle == null) ? EMPTY : headerStyle, dataGetter, dataStyleFunc);
    }
}
