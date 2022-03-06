package com.alibaba.easyexcel.course.task4.test;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import com.alibaba.easyexcel.course.base.model.Demo;
import com.alibaba.easyexcel.course.base.model.LargeData;
import com.alibaba.easyexcel.course.base.utils.FileUtils;
import com.alibaba.easyexcel.course.task3.MyExcelWriter;
import com.alibaba.easyexcel.course.task3.MyExcelWriterImpl;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

/**
 * 封装一个MyExcelWriter
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class Task3_03_MyExcelWriter {

    /**
     * 基于前面的代码手写一个MyExcelWriter工具类
     * 任务说明：
     * 我已经会了如果导出一个xlsx以及如何通过cglib来反射读取到一个对象成员变量以及值，这个时候 我们可以结合这2个知识点，来完成一个我们自己的excel导出的工具类：MyExcelWriter
     * 需要完成：
     * 1. MyExcelWriterImpl里面 根据 fieldList的数据 写入头数数据
     * 2. MyExcelWriterImpl.write 方法
     *
     * @throws Exception
     */
    @Test
    public void write() throws Exception {
        // 因为我们 的 MyExcelWriter 实现了AutoCloseable， 所以可以用 tray with resource 来自动关闭 MyExcelWriter
        try (MyExcelWriter<Demo> myExcelWriter = new MyExcelWriterImpl<>(
            FileUtils.getPath() + "write/demo" + System.currentTimeMillis() + ".xlsx", Demo.class)) {
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");

            // 这里 的数据 正式使用中 一般是数据库查询出来
            List<Demo> demoList = new ArrayList<>();
            for (int i = 1; i <= 10; i++) {
                Demo demo = new Demo();
                demo.setString("字符串" + i);
                demo.setDate(simpleDateFormat.parse("2022-01_" + i));
                demo.setInteger(i);
                demoList.add(demo);
            }

            // 写入到excel
            myExcelWriter.write(demoList);
        }
        log.info("完成文件输出");
    }

    /**
     * 用poi 使用XSSFWorkbook 输出一个大文件
     * 任务说明：
     * 使用poi的XSSFWorkbook 来输出一个xlsx,可以参照官方文档：https://poi.apache.org/components/spreadsheet/quick-guide.html
     * 或者 自己搜索 poi XSSFWorkbook 写入
     * 需要完成：
     * 1. 写入10W行25列数据，25列数据格式为： 字符串_(行号)_(列号)
     * 2. 设置运行内存为128M -Xmx128M  再次运行 导出，不会设置的同学可以搜索： idea(如果其实其他编辑器，则搜索对应的编辑器) junit 设置最大内存
     *
     * @throws Exception
     */
    @Test
    public void poiHssfLargeWrite() throws Exception {
        // TODO
        // 1. 将最大行数设置为10W行 运行 并记录时间
        // 2. 设置最大运行内存为128M  -Xmx128M 直接抛异常

        // 最大行数
        int maxColumn = 10 * 10000;
        int logLine = maxColumn / 10;

        long start = System.currentTimeMillis();

        String fileName = FileUtils.getPath() + "write/large" + System.currentTimeMillis() + ".xlsx";
        new File(fileName).getParentFile().mkdirs();
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");
        for (int i = 0; i < maxColumn; i++) {
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < 25; j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue("字符串" + i + "_" + j);
            }
            if (i % logLine == 0) {
                log.info("第{}行输出完成", i);
            }
        }
        workbook.write(new FileOutputStream(fileName));
        workbook.close();
        log.info("完成文件输出,耗时:{}", System.currentTimeMillis() - start);
    }

    /**
     * 用poi 使用SXSSFWorkbook 输出一个大文件
     * 任务说明：
     * 使用poi的SXSSFWorkbook 来输出一个xlsx,可以参照官方文档：https://poi.apache.org/components/spreadsheet/quick-guide.html
     * 或者 自己搜索 poi SXSSFWorkbook 写入
     * 需要完成：
     * 1. 写入10W行25列数据，25列数据格式为： 字符串_(行号)_(列号)
     * 2. 设置运行内存为128M -Xmx128M  再次运行 导出
     *
     * @throws Exception
     */
    @Test
    public void poiSxssfLargeWrite() throws Exception {
        // TODO
        // 1. 将最大行数设置为10W行 运行 并记录时间
        // 2. 设置最大运行内存为128M  -Xmx128M 正常运行 耗时几乎不变

        // 最大行数
        int maxColumn = 10 * 10000;
        int logLine = maxColumn / 10;

        long start = System.currentTimeMillis();

        String fileName = FileUtils.getPath() + "write/large" + System.currentTimeMillis() + ".xlsx";
        new File(fileName).getParentFile().mkdirs();
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        SXSSFSheet sheet = workbook.createSheet("Sheet1");
        for (int i = 0; i < maxColumn; i++) {
            SXSSFRow row = sheet.createRow(i);
            for (int j = 0; j < 25; j++) {
                SXSSFCell cell = row.createCell(j);
                cell.setCellValue("字符串" + j + "_" + i);
            }
            if (i % logLine == 0) {
                log.info("第{}行输出完成", i);
            }
        }
        workbook.write(new FileOutputStream(fileName));
        // SXSSF 模式 需要额外调用dispose
        workbook.dispose();
        workbook.close();
        log.info("完成文件输出,耗时:{}", System.currentTimeMillis() - start);
    }

    /**
     * 用MyExcelWriter 输出一个大文件
     * 任务说明：
     * 用于对比poi，直接参照：write 第一个测试案例来输出，写入的对象 直接用LargeData即可
     * 需要完成：
     * 1. 写入10W行25列数据，25列数据格式为： 字符串_(行号)_(列号)
     * 2. 设置运行内存为128M -Xmx128M  再次运行 导出
     * 3. 用128M 内存导出100W数据
     *
     * @throws Exception
     */
    @Test
    public void myExcelWriterLargeWrite() throws Exception {
        // 1. 将最大行数设置为10W行 运行 并记录时间
        // 2. 设置最大运行内存为128M  -Xmx128M 正常运行 耗时几乎不变
        // 3. 用128M 导出100W 数据

        // 最大行数
        int maxColumn = 10 * 10000;
        int logLine = maxColumn / 10;

        long start = System.currentTimeMillis();
        String fileName = FileUtils.getPath() + "write/large" + System.currentTimeMillis() + ".xlsx";

        // 用tray with resource 自动关闭 MyExcelWriter
        try (MyExcelWriter<LargeData> myExcelWriter = new MyExcelWriterImpl<>(fileName, LargeData.class)) {

            for (int i = 0; i < maxColumn; i++) {
                LargeData largeData = largeData(i);
                List<LargeData> list = new ArrayList<>();
                list.add(largeData);
                // 写入到excel
                myExcelWriter.write(list);
                if (i % logLine == 0) {
                    log.info("第{}行输出完成", i);
                }
            }
        }
        log.info("完成文件输出,耗时:{}", System.currentTimeMillis() - start);
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
