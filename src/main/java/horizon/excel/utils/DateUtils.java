package horizon.excel.utils;

import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;

/**
 * @author K·J
 * <p>
 * Create at 2018-06-14 11:47
 */
public class DateUtils {
    /**
     * java.util.Date 转为 java.time.LocalDateTime
     *
     * @param date Date
     * @return LocalDateTime
     */
    public static LocalDateTime dateToLocalDateTime(Date date) {
        if (null == date) {
            return null;
        }
        Instant instant = date.toInstant();
        ZoneId zone = ZoneId.systemDefault();
        return LocalDateTime.ofInstant(instant, zone);
    }

    /**
     * java.util.Date 转为 java.time.LocalDate
     *
     * @param date Date
     * @return LocalDate
     */
    public static LocalDate dateToLocalDate(Date date) {
        if (null == date) {
            return null;
        }
        Instant instant = date.toInstant();
        ZoneId zone = ZoneId.systemDefault();
        LocalDateTime localDateTime = LocalDateTime.ofInstant(instant, zone);
        return localDateTime.toLocalDate();
    }

    /**
     * java.time.LocalDateTime 转为 java.util.Date
     *
     * @param localDateTime LocalDateTime
     * @return Date
     */
    public static Date localDateTimeToDate(LocalDateTime localDateTime) {
        ZoneId zone = ZoneId.systemDefault();
        Instant instant = localDateTime.atZone(zone).toInstant();
        return Date.from(instant);
    }
}
