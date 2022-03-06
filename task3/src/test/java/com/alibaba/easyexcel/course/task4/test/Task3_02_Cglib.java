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
 * 了解什么是cglib
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class Task3_02_Cglib {

    /**
     * 通过反射 读取值
     * 任务说明：反射来读取一个变量的值在java中很常见，但是性能不是很好，我们可以先尝试用反射读取一个变量的值。
     * 需要完成：
     * 1. 日志输入一个对象下面所有成员变量的值
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

        // TODO 这里需要完成
        // 通过反射来输入 上面2个对象的成员变量以及值
        reflectPrintData(demo);
        reflectPrintData(demo2);
    }

    /**
     * 通过反射的方法 输出一个对象里面所有的值
     *
     * @param obj
     */
    private void reflectPrintData(Object obj) throws Exception {
        // TODO
        // 可以先拿到class  然后 通过 clazz.getDeclaredFields来拿到所有成员变量的field
        // 然后 自己拼接 getString() 等方法 ，调用 clazz.getMethod(method).invoke(obj, null) 拿到成员变量的值
        // 最后输出的日志 如下
        //  log.info("成员变量:{}的值为:{}", ,);

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

    /**
     * 通过cglib 读取值
     * 任务说明：可以自己搜索 cglib  beanmap 关键字来 查看cglib的用法
     * 需要完成：
     * 1. 使用日志输入一个对象下面所有成员变量的值
     *
     * @throws Exception
     */
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

        // TODO
        // 使用cglib的BeanMap 来输出
        cglibPrintData(demo);
        cglibPrintData(demo2);
    }

    /**
     * 通过cglib 一个对象里面所有的值一个对象下面所有成员变量的值
     *
     * @param obj
     */
    private void cglibPrintData(Object obj) throws Exception {
        // TODO
        //  BeanMap.create 可以创建一个 Map<成员变量的名字,成员变量的值>
        // 根据 map.get(key) 拿到成员变量的值
        //  log.info("成员变量:{}的值为:{}", ,);
        BeanMap beanMap = BeanMap.create(obj);
        for (Object key : beanMap.keySet()) {
            log.info("成员变量:{}的值为:{}", key, beanMap.get(key));
        }
    }

    /**
     * 对比反射和cglib的性能
     * 任务说明：生成1000W 个 Demo，确保里面的 string date integer 都是不一样的值，然后通过反射 以及 cglib 2个种方法实现，对比他们2个的性能
     * 需要完成：
     * 1. 完成反射读取对象中的值
     * 2. 完成Cglib读取对象中的值
     *
     * @throws Exception
     */
    @Test
    public void reflectVsCglib() throws Exception {
        // 运行的次数
        int count = 1000 * 10000;
        long start = System.currentTimeMillis();
        // 反射
        for (int i = 0; i < count; i++) {
            Demo demo = new Demo();
            demo.setString("字符串" + i);
            demo.setDate(new Date());
            demo.setInteger(i);

            // TODO 反射读取对象中的值
            // 这里别调用reflectPrintData 不然不会输出一堆对象
            // 需要调用 clazz.getMethod(method).invoke(demo, null);方法 拿值 可以不输出
            // 拿到对应的class
            Class<?> clazz = demo.getClass();
            // 拿到所有的属性
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                String fieldName = field.getName();
                // 生成get方法
                String method = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                clazz.getMethod(method).invoke(demo, null);
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

            // TODO Cglib读取对象中的值
            // 这里别调用cglibPrintData 不然不会输出一堆对象
            // 需要调用  beanMap.get(key);方法 拿值 可以不输出
            BeanMap beanMap = BeanMap.create(demo);
            for (Object key : beanMap.keySet()) {
                beanMap.get(key);
            }
        }
        log.info("cglib耗时：{}", System.currentTimeMillis() - start);
    }

    /**
     * 查看cglib生成的class源码
     * 任务说明：
     * 设置cglib输出，看下cblib最终输出的class
     * 需要完成：
     * 1. 查看生成的代理类的get 方法
     *
     * @throws Exception
     */
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
