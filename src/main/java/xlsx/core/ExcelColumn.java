package xlsx.core;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

import java.util.function.Function;

/**
 * ### dataGetter ###<p>
 * Value from this getter will be processed by auto-magic.
 * <p> Any Date classes will be transformer into {@link java.util.Calendar}
 * <p> Enum will write enum value name. Example: UserRole.ADMIN in excel will be ADMIN as String
 *
 * @author Daniils Loputevs
 */
@Setter
@Getter
@AllArgsConstructor
public class ExcelColumn<D> {
    private String headerValue;
    private ExcelCellStyle headerStyle;
    private Function<D, Object> dataGetter;
    private Function<D, ExcelCellStyle> dataStyle;
}
