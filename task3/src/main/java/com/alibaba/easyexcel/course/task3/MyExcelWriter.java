package com.alibaba.easyexcel.course.task3;

import java.util.List;

/**
 * 自定义的excel 写入的类
 *
 * @param <T>
 */
public interface MyExcelWriter<T> extends AutoCloseable {

    /**
     * 往excel 中写入数据
     *
     * @param dataList
     * @throws Exception
     */
    void write(List<T> dataList) throws Exception;

    /**
     * 完成excel 写入的方法
     *
     * @throws Exception
     */
    void finish() throws Exception;
}
