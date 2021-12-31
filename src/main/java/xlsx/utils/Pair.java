package xlsx.utils;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @author Daniils Loputevs
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class Pair<A, B> {
    private A first;
    private B second;
}
