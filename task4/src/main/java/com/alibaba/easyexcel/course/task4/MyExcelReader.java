package com.alibaba.easyexcel.course.task4;

import java.util.List;
import java.util.function.Consumer;

/**
 * 读取excel的工具
 *
 * @param <T>
 * @author Jiaju Zhuang
 */
public interface MyExcelReader<T> {
    /**
     * 读取一个excel
     *
     * @param file     需要读取的excel文件
     * @param clazz    对应的class 类
     * @param consumer 读取到excel一行数据的回调方法。 这个是jdk8 以后的函数式接口，不懂的同学可以搜索下： Consumer jdk8 去普及下知识
     */
    void read(String file, Class<T> clazz, Consumer<List<T>> consumer);
}
