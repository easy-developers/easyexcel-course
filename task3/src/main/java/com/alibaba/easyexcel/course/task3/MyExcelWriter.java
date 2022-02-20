package com.alibaba.easyexcel.course.task3;

import java.util.List;

public interface MyExcelWriter<T> extends AutoCloseable{

    void write(List<T> dataList) throws Exception;

    void finish() throws Exception;

}
