package xlsx.utils;

import lombok.val;
import models.User;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Random;

import static xlsx.utils.DateUtil.toLocalDateTime;

public class RandomTestDataGenerator {
    
    public static int randBetween(int start, int end) {
        return start + (int) Math.round(Math.random() * (end - start));
    }
    
    public Iterable<User> genRandomUsers(int count) {
        val rsl = new ArrayList<User>();
        val random = new Random();
        
        for (int i = 0; i <= count; i++) {
            rsl.add(new User(
                    (long) i * randBetween(1, 9),
                    "name", // TODO : random string
                    random.nextBoolean() ? User.Role.USER : User.Role.ADMIN,
                    toLocalDateTime(randomCalendar()),
                    random.nextBoolean(),
                    BigDecimal.valueOf(random.nextDouble())
            ));
        }
        return rsl;
        
    }
    
    private Calendar randomCalendar() {
        val calendar = new GregorianCalendar();
        int year = randBetween(1900, 2010);
        calendar.set(calendar.YEAR, year);
        int dayOfYear = randBetween(1, calendar.getActualMaximum(calendar.DAY_OF_YEAR));
        calendar.set(calendar.DAY_OF_YEAR, dayOfYear);
//        System.out.println(calendar.get(calendar.YEAR) + "-" + (calendar.get(calendar.MONTH) + 1) + "-" + calendar.get(calendar.DAY_OF_MONTH));
        return calendar;
    }
}
