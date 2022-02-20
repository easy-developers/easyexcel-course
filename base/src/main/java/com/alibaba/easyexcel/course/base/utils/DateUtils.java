package com.alibaba.easyexcel.course.base.utils;

import java.time.Duration;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.util.Date;

import lombok.extern.slf4j.Slf4j;

/**
 * 日期工具类
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class DateUtils {

    /**
     * 一天毫秒数
     */
    private static final long DAY_MILLISECONDS = 86400000L;
    /**
     * 1900年的日期
     */
    private static final LocalDateTime LOCAL_DATE_TIME_1900 = LocalDateTime.of(1900, 1, 1, 0, 0, 0, 0);

    /**
     * 将一个excel里面的日期字符串转成java 日期
     *
     * @param stringValue
     * @return
     */
    public static Date convertToJavaDate(String stringValue) {
        if (stringValue == null) {
            return null;
        }
        double value = Double.parseDouble(stringValue);
        int wholeDays = (int)Math.floor(value);
        long millisecondsInDay = (long)(((value - wholeDays) * DAY_MILLISECONDS) + 0.5D);

        int startYear = 1900;
        int dayAdjust = -1;
        if (wholeDays < 61) {
            dayAdjust = 0;
        }
        LocalDate localDate = LocalDate.of(startYear, 1, 1).plusDays((long)wholeDays + dayAdjust - 1);
        LocalTime localTime = LocalTime.ofNanoOfDay(millisecondsInDay * 1000000);
        LocalDateTime localDateTime = LocalDateTime.of(localDate, localTime);

        ZoneId zone = ZoneId.systemDefault();
        Instant instant = localDateTime.atZone(zone).toInstant();
        return Date.from(instant);
    }

    /**
     * 将一个java的日期转成excel的日期
     *
     * @param date
     * @return
     */
    public static Double convertToExcelDate(Date date) {
        if (date == null) {
            return null;
        }
        LocalDateTime localDateTime = LocalDateTime.ofInstant(Instant.ofEpochMilli(date.getTime()),
            ZoneId.systemDefault());
        Duration duration = Duration.between(LOCAL_DATE_TIME_1900, localDateTime);
        return duration.getSeconds() / 86400.0 + duration.getNano() / 86400000000000.0 + 2;
    }


}
