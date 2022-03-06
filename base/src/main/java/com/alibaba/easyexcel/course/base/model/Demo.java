package com.alibaba.easyexcel.course.base.model;

import java.util.Date;

import lombok.Data;
import lombok.ToString;

/**
 * demo对象
 */
@Data
@ToString
public class Demo {
    /**
     * 字符串
     */
    private String string;
    /**
     * 日期
     */
    private Date date;
    /**
     * 数字
     */
    private Integer integer;
}
