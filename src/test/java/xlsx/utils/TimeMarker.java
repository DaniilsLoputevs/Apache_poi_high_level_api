package xlsx.utils;

import lombok.val;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.LinkedHashMap;
import java.util.concurrent.TimeUnit;

public class TimeMarker {
    private static final LinkedHashMap<String, Calendar> times = new LinkedHashMap<>();

    public static void setMark(String markName) {
        times.put(markName, new GregorianCalendar());
        DateUtil.print(markName, new GregorianCalendar());
    }

    public static void printState() {
        boolean first = true;
        Calendar prevCalendar = null;
        for (val t : times.entrySet()) {
            if (first) {
                first = false;
                DateUtil.print(t.getKey(), t.getValue());

            } else {
//                DateUtils.print(t.getKey(), t.getValue());

                val millis = t.getValue().getTimeInMillis() - prevCalendar.getTimeInMillis();

                val formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                System.out.println(t.getKey() + " : " + formatter.format(t.getValue().getTime())
                        + " time after last mark = " + TimeUnit.MILLISECONDS.toMillis(millis));
            }
            prevCalendar = t.getValue();
        }
    }


}
