package com.alibaba.easyexcel.course.task4.test;

import java.lang.reflect.Field;
import java.util.Date;

import com.alibaba.easyexcel.course.base.model.Demo;
import com.alibaba.easyexcel.course.base.model.Demo2;
import com.alibaba.easyexcel.course.base.utils.FileUtils;

import lombok.extern.slf4j.Slf4j;
import net.sf.cglib.beans.BeanMap;
import net.sf.cglib.core.DebuggingClassWriter;
import org.junit.Test;

/**
 * 任务目标： 读取中xlsx中的信息
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class Task3_02_Cglib {

    /**
     * 了解反射
     * 书写一个
     *
     * @throws Exception
     */
    @Test
    public void reflect() throws Exception {
        Demo demo = new Demo();
        demo.setString("字符串1");
        demo.setDate(new Date());
        demo.setInteger(1);

        Demo2 demo2 = new Demo2();
        demo2.setString("字符串2");
        demo2.setDate(new Date());
        demo2.setInteger(2);
        reflectPrintData(demo);
        reflectPrintData(demo2);
    }

    /**
     * 通过反射的方法 一个对象里面所有的值
     *
     * @param obj
     */
    private void reflectPrintData(Object obj) throws Exception {
        // 拿到对应的class
        Class<?> clazz = obj.getClass();
        // 拿到所有的属性
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            String fieldName = field.getName();
            // 生成get方法
            String method = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
            Object data = clazz.getMethod(method).invoke(obj, null);
            log.info("成员变量:{}的值为:{}", field.getName(), data);
        }
    }

    @Test
    public void cglib() throws Exception {
        Demo demo = new Demo();
        demo.setString("字符串1");
        demo.setDate(new Date());
        demo.setInteger(1);

        Demo2 demo2 = new Demo2();
        demo2.setString("字符串2");
        demo2.setDate(new Date());
        demo2.setInteger(2);

        cglibPrintData(demo);
        cglibPrintData(demo2);
    }

    /**
     * 通过cglib 一个对象里面所有的值
     *
     * @param obj
     */
    private void cglibPrintData(Object obj) throws Exception {
        BeanMap beanMap = BeanMap.create(obj);
        for (Object key : beanMap.keySet()) {
            log.info("成员变量:{}的值为:{}", key, beanMap.get(key));
        }
    }

    @Test
    public void reflectVsCglib() throws Exception {
        int count = 100 * 10000;
        long start = System.currentTimeMillis();
        // 反射
        for (int i = 0; i < count; i++) {
            Demo demo = new Demo();
            demo.setString("字符串" + i);
            demo.setDate(new Date());
            demo.setInteger(i);

            // 拿到对应的class
            Class<?> clazz = demo.getClass();
            // 拿到所有的属性
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                String fieldName = field.getName();
                // 生成get方法
                String method = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                Object data = clazz.getMethod(method).invoke(demo, null);
                if (i == 0) {
                    log.info("成员变量:{}的值为:{}", field.getName(), data);
                }
            }
        }
        log.info("反射耗时：{}", System.currentTimeMillis() - start);

        start = System.currentTimeMillis();
        // cglib
        for (int i = 0; i < count; i++) {
            Demo demo = new Demo();
            demo.setString("字符串" + i);
            demo.setDate(new Date());
            demo.setInteger(i);

            BeanMap beanMap = BeanMap.create(demo);
            for (Object key : beanMap.keySet()) {
                Object data = beanMap.get(key);
                if (i == 0) {
                    log.info("成员变量:{}的值为:{}", key, data);
                }
            }
        }
        log.info("cglib耗时：{}", System.currentTimeMillis() - start);
    }

    @Test
    public void cglibAnalyse() throws Exception {
        // 输出cglib 生成的class
        System.getProperties().put("sun.misc.ProxyGenerator.saveGeneratedFiles", "true");
        // 指定输出的位置
        // 输出的目录 就是 easyexcel-course/task3/target/test-classes/cglib/com/alibaba/easyexcel/course/task3/
        System.setProperty(DebuggingClassWriter.DEBUG_LOCATION_PROPERTY, FileUtils.getPath() + "cglib");

        Demo demo = new Demo();
        demo.setString("字符串1");
        demo.setDate(new Date());
        demo.setInteger(1);


        BeanMap beanMap = BeanMap.create(demo);
        for (Object key : beanMap.keySet()) {
            Object data = beanMap.get(key);
            log.info("成员变量:{}的值为:{}", key, data);
        }
    }
}
