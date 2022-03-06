package com.alibaba.easyexcel.course.task4.test;

import java.lang.reflect.Field;
import java.util.Date;

import com.alibaba.easyexcel.course.base.model.Demo;
import com.alibaba.easyexcel.course.base.model.Demo2;
import com.alibaba.easyexcel.course.base.utils.FileUtils;

import lombok.extern.slf4j.Slf4j;
import net.sf.cglib.beans.BeanMap;
import net.sf.cglib.core.DebuggingClassWriter;
import org.apache.poi.ss.formula.functions.T;
import org.junit.Test;

/**
 * cglib的写
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class Task4_02_Cglib {

    /**
     * 通过cglib 写入值
     * 任务说明： 熟悉cglib 如何往一个对象写入值
     * 需要完成：
     * 1. 阅读cglib往对象写入数据的代码
     *
     * @throws Exception
     */
    @Test
    public void cglib() throws Exception {
        // 假设我么要构建一个 Demo 对象
        Demo demo = cglibBuildData(Demo.class, "字符串1", new Date(), 1);
        log.info("构建完成的对象是:{}", demo);
    }

    /**
     * 构建一个对象放入值
     *
     * @param clazz
     * @param data  字段1-3 的值
     * @throws Exception
     */
    private <T> T cglibBuildData(Class<T> clazz, Object... data) throws Exception {
        // 实例化一个对象
        T obj = clazz.newInstance();
        // 拿到对应class 所有的成员变量
        Field[] fieldList = clazz.getDeclaredFields();
        // cglib 构建bean map
        BeanMap beanMap = BeanMap.create(obj);
        for (int i = 0; i < fieldList.length; i++) {
            // 放入值
            beanMap.put(fieldList[i].getName(), data[i]);
        }
        return obj;
    }

    /**
     * 查看cglib生成的class源码
     * 任务说明：
     * 设置cglib输出，看下cblib最终输出的class
     * 需要完成：
     * 1. 查看生成的代理类的put 方法 源码基本和get一致
     *
     * @throws Exception
     */
    @Test
    public void cglibAnalyse() throws Exception {
        // 输出cglib 生成的class
        System.getProperties().put("sun.misc.ProxyGenerator.saveGeneratedFiles", "true");
        // 指定输出的位置
        // 输出的目录 就是 easyexcel-course/task4/target/test-classes/cglib/com/alibaba/easyexcel/course/task3/
        System.setProperty(DebuggingClassWriter.DEBUG_LOCATION_PROPERTY, FileUtils.getPath() + "cglib");

        Demo demo = cglibBuildData(Demo.class, "字符串1", new Date(), 1);
        log.info("构建完成的对象是:{}", demo);
    }
}
