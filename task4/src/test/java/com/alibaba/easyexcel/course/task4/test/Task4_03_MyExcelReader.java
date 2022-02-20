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
 * 任务目标： 读取中xlsx中的信息
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class Task4_03_MyExcelReader {

    /**
     * 了解反射
     * 书写一个
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
     *
     * @throws Exception
     */
    @Test
    public void poiLargeRead() throws Exception {
        // 1. 将最大行数设置为10W行 运行 并记录时间
        // 2. 设置最大运行内存为128M -Xmx128M 直接报错

        // 复制写的代码过来

        // 最大行数
        int maxColumn = 10 * 10000;
        int logLine = maxColumn / 10;

        long start = System.currentTimeMillis();

        String fileName = FileUtils.getPath() + "read/large" + System.currentTimeMillis() + ".xlsx";
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

        //  读取excel poi读 xlsx 读值有HSSF一个模式
        start = System.currentTimeMillis();
        XSSFWorkbook readWorkbook = new XSSFWorkbook(fileName);
        XSSFSheet readSheet = readWorkbook.getSheetAt(0);
        int lastRowNum = readSheet.getLastRowNum();

        // 临时存放数据 模拟存储数据库
        List<List<String>> tempData = new ArrayList<>();
        for (int i = 0; i <= lastRowNum; i++) {
            XSSFRow row = readSheet.getRow(i);
            List<String> rowData = new ArrayList<>();
            tempData.add(rowData);
            int lastCellNum = row.getLastCellNum();
            for (int j = 0; j < lastCellNum; j++) {
                rowData.add(row.getCell(j).getStringCellValue());
            }
            if (i % logLine == 0) {
                log.info("第{}行读取完成,数据为{}", i, JSON.toJSONString(rowData));
            }
            // 模拟每隔100条 存储数据库
            if (i % 100 == 0) {
                tempData = new ArrayList<>();
            }
        }

        log.info("完成文件读取,耗时:{}", System.currentTimeMillis() - start);

    }

    /**
     * 用MyExcelReader 读取一个大文件
     *
     * @throws Exception
     */
    @Test
    public void myExcelReaderLargeRead() throws Exception {
        // 1. 将最大行数设置为10W行 运行 并记录时间
        // 2. 设置最大运行内存为128M -Xmx128M 正常运行

        // 复制写的代码过来
        // 最大行数
        int maxColumn = 10 * 10000;
        int logLine = maxColumn / 10;

        long start = System.currentTimeMillis();
        String fileName = FileUtils.getPath() + "reade/large" + System.currentTimeMillis() + ".xlsx";
        new File(fileName).getParentFile().mkdirs();

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

        // 读excel
        start = System.currentTimeMillis();
        MyExcelReader<LargeData> myExcelReader = new MyExcelReaderImpl<>();
        AtomicInteger count = new AtomicInteger();
        myExcelReader.read(fileName, LargeData.class, list -> {
            for (LargeData largeData : list) {
                count.incrementAndGet();
                if (count.get() % logLine == 0) {
                    log.info("第{}行输出完成,数据为{}", count.get(), JSON.toJSONString(largeData));
                }
            }
        });
        log.info("完成文件读取,耗时:{}", System.currentTimeMillis() - start);
    }

    /**
     * 用MyExcelReader 读取一个大文件
     *
     * @throws Exception
     */
    @Test
    public void myExcelReaderSaxLargeRead() throws Exception {
        // 1. 将最大行数设置为10W行 运行 并记录时间
        // 2. 设置最大运行内存为128M -Xmx128M 正常运行

        // 复制写的代码过来
        // 最大行数
        int maxColumn = 10 * 10000;
        int logLine = maxColumn / 10;

        long start = System.currentTimeMillis();
        String fileName = FileUtils.getPath() + "reade/large" + System.currentTimeMillis() + ".xlsx";
        new File(fileName).getParentFile().mkdirs();

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

        // 读excel
        start = System.currentTimeMillis();
        MyExcelReader<LargeData> myExcelReader = new MyExcelReaderSaxImpl<>();
        AtomicInteger count = new AtomicInteger();
        myExcelReader.read(fileName, LargeData.class, list -> {
            for (LargeData largeData : list) {
                count.incrementAndGet();
                if (count.get() % logLine == 0) {
                    log.info("第{}行输出完成,数据为{}", count.get(), JSON.toJSONString(largeData));
                }
            }
        });
        log.info("完成文件读取,耗时:{}", System.currentTimeMillis() - start);
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
