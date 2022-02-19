package com.alibaba.easyexcel.course.base.utils;

/**
 * 抛出异常的 Consumer
 *
 * @param <T>
 */
@FunctionalInterface
interface ConsumerThrowsException<T> {

    void accept(T t) throws Exception;
}
