package com.alibaba.easyexcel.course.task4;

import java.util.List;
import java.util.function.Consumer;

public interface MyExcelReader<T> {
    void read(String file, Class<T> clazz, Consumer<List<T>> consumer);
}
