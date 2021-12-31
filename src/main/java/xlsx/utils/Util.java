package xlsx.utils;

import java.util.Collection;

public class Util {

    public static int sizeOfIterable(Iterable<?> iterable) {
        if (iterable instanceof Collection<?>)
            return ((Collection<?>) iterable).size();
        
        int counter = 0;
        for (Object ignored : iterable) {
            counter++;
        }
        return counter;
    }
}
