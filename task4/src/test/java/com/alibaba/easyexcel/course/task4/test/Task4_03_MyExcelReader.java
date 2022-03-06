package com.alibaba.easyexcel.course.task4.test;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

import com.alibaba.easyexcel.course.base.model.Demo;
import com.alibaba.easyexcel.course.base.model.LargeData;
import com.alibaba.easyexcel.course.base.utils.FileUtils;
import com.alibaba.easyexcel.course.task3.MyExcelWriter;
import com.alibaba.easyexcel.course.task3.MyExcelWriterImpl;
import com.alibaba.easyexcel.course.task4.MyExcelReader;
import com.alibaba.easyexcel.course.task4.MyExcelReaderImpl;
import com.alibaba.easyexcel.course.task4.MyExcelReaderSaxImpl;
import com.alibaba.fastjson.JSON;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

/**
 * 封装一个MyExcelReader
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class Task4_03_MyExcelReader {

    /**
     * ●基于前面的代码手写一个MyExcelReader工具类
     * 任务说明：
     * 我已经会了如果读取一个xlsx以，结合前面写的时候需要用一个对象去写，来完成一个我们自己的excel读取的工具类：MyExcelReader
     * 需要完成：
     * 1. MyExcelReaderImpl.readSharedStrings 读取readSharedStrings
     * 2. MyExcelReaderImpl.readSheet1 读取sheet 并回调给 consumer
     *
     * @throws Exception
     */
    @Test
    public void read() throws Exception {
        MyExcelReader<Demo> myExcelReader = new MyExcelReaderImpl<>();
        myExcelReader.read(FileUtils.getPath() + "demo.xlsx", Demo.class, list -> {
            // 这里我们在工作中要做的 就是去数据库存储
            log.info("读取到{}条数据", list.size());
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            for (Demo demo : list) {
                log.info("数据:string:{},date:{},integer:{}", demo.getString(), simpleDateFormat.format(demo.getDate()),
                    demo.getInteger());
            }
        });
    }

    /**
     * 用poi 读取一个大文件
     * 任务说明：
     * 使用poi的XSSFWorkbook 来读取一个xlsx,可以参照官方文档：https://poi.apache.org/components/spreadsheet/quick-guide.html
     * 或者 自己搜索 poi XSSFWorkbook 读取
     * 需要完成：
     * 1. 写入10W行25列数据，每一个单元格的数据格式为： 字符串_(行号)_(列号) 并记录耗时，然后读取这个excel, 并记录耗时，这里注意需要模拟每隔100条，清空数据，和我们的MyExcelReader一致
     * 2. 设置运行内存为128M -Xmx128M ，会发现内存溢出了，因为poi读取是在内存中
     *
     * @throws Exception
     */
    @Test
    public void poiLargeRead() throws Exception {
        // 1. 将最大行数设置为10W行 运行 并记录时间
        // 2. 设置最大运行内存为128M -Xmx128M 直接报错

        // 复制写的代码过来

    }

    /**
     * 用MyExcelWriter 读取一个大文件
     * 任务说明：
     * 用于对比poi，直接参照：read 第一个测试案例来输出，写入的对象 直接用LargeData即可
     * 需要完成：
     * 1. 写入10W行25列数据，每一个单元格的数据格式为： 字符串_(行号)_(列号) 并记录耗时, 然后读取这个数据
     * 2. 设置运行内存为128M -Xmx128M  再次运行
     * 3. 用128M 内存导出100W数据
     *
     * @throws Exception
     */
    @Test
    public void myExcelReaderLargeRead() throws Exception {
        // 1. 将最大行数设置为10W行 运行 并记录时间
        // 2. 设置最大运行内存为128M -Xmx128M 正常运行
        // 3. 修改行数为100W 查看结果

        // 复制写的代码过来
    }

    /**
     * 用MyExcelWriterSax 读取一个大文件
     * 任务说明：
     * 由于dom4j读取sheet1.xml 还是在内存里面，在遇到100W量级的数据还是会OOM，这个时候我们需要引入：SAXParserFactory 这个会一行一行的读取 xml中的数据
     * 需要完成：
     * 1. 由于sax读取相对比较复杂，所以不再让大家书写了。请阅读 并理解 MyExcelReaderSaxImpl 的代码
     * 1. 写入10W行25列数据，每一个单元格的数据格式为： 字符串_(行号)_(列号) 并记录耗时, 然后读取这个数据
     * 2. 设置运行内存为128M -Xmx128M  再次运行
     * 3. 用128M 内存导出并读取100W数据
     *
     * @throws Exception
     */
    @Test
    public void myExcelReaderSaxLargeRead() throws Exception {
        // 1. 将最大行数设置为10W行 运行 并记录时间
        // 2. 设置最大运行内存为128M -Xmx128M 正常运行
        // 3. 用128M 内存导出并读取100W数据

        // 复制写的代码过来
    }

    private LargeData largeData(int row) {
        LargeData largeData = new LargeData();
        largeData.setStr1("字符串1_" + row);
        largeData.setStr2("字符串2_" + row);
        largeData.setStr3("字符串3_" + row);
        largeData.setStr4("字符串4_" + row);
        largeData.setStr5("字符串5_" + row);
        largeData.setStr6("字符串6_" + row);
        largeData.setStr7("字符串7_" + row);
        largeData.setStr8("字符串8_" + row);
        largeData.setStr9("字符串9_" + row);
        largeData.setStr10("字符串10_" + row);
        largeData.setStr11("字符串11_" + row);
        largeData.setStr12("字符串12_" + row);
        largeData.setStr13("字符串13_" + row);
        largeData.setStr14("字符串14_" + row);
        largeData.setStr15("字符串15_" + row);
        largeData.setStr16("字符串16_" + row);
        largeData.setStr17("字符串17_" + row);
        largeData.setStr18("字符串18_" + row);
        largeData.setStr19("字符串19_" + row);
        largeData.setStr20("字符串20_" + row);
        largeData.setStr21("字符串21_" + row);
        largeData.setStr22("字符串22_" + row);
        largeData.setStr23("字符串23_" + row);
        largeData.setStr24("字符串24_" + row);
        largeData.setStr25("字符串25_" + row);
        return largeData;
    }

}
