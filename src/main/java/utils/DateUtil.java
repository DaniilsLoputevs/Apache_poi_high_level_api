package utils;

import lombok.val;

import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

public class DateUtil {
    
    // для большего перевода одних блин дат в другие блин даты: https://www.logicbig.com/how-to/java-8-date-time-api/to-date-conversion.html
    
    public static Calendar toCalendar(LocalDateTime localDateTime) {
        return GregorianCalendar.from(ZonedDateTime.of(localDateTime, ZoneId.systemDefault()));
    }
    
    public static Calendar toCalendar(LocalDate localDate) {
        Calendar calendar = Calendar.getInstance();
        calendar.clear();
        //assuming start of day
        calendar.set(localDate.getYear(), localDate.getMonthValue() - 1, localDate.getDayOfMonth());
        return calendar;
    }
    
    public static Calendar toCalendar(Date date) {
        val cal = Calendar.getInstance();
        cal.setTime(date);
        return cal;
    }
    
    // print hepler
    
    public static LocalDateTime toLocalDateTime(Calendar calendar) {
        return LocalDateTime.ofInstant(calendar.toInstant(), calendar.getTimeZone().toZoneId());
    }
    
    
    private static void printCalendar(String name, Calendar calendar) {
        printCalendar(name, calendar, "yyyy-MM-dd HH:mm:ss");
    }
    
    private static void printLocalDate(String name, LocalDate localDate) {
        printLocalDate(name, localDate, "yyyy-MM-dd");
    }
    
    private static void printLocalDateTime(String name, LocalDateTime localDateTime) {
        printLocalDateTime(name, localDateTime, "yyyy-MM-dd HH:mm:ss");
    }
    
    private static void printCalendar(String name, Calendar calendar, String pattern) {
        val format = new SimpleDateFormat(pattern);
        System.out.println(name + " : " + format.format(calendar.getTime()));
    }
    
    private static void printLocalDate(String name, LocalDate localDate, String pattern) {
        val format = DateTimeFormatter.ofPattern(pattern);
        System.out.println(name + " : " + localDate.format(format));
    }
    
    private static void printLocalDateTime(String name, LocalDateTime localDateTime, String pattern) {
        val format = DateTimeFormatter.ofPattern(pattern);
        System.out.println(name + " : " + localDateTime.format(format));
    }
}
