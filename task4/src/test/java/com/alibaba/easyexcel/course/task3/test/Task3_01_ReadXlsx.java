package com.alibaba.easyexcel.course.task3.test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import com.alibaba.easyexcel.course.base.utils.FileUtils;

import lombok.extern.slf4j.Slf4j;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.junit.Test;

/**
 * 任务目标： 读取中xlsx中的信息
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class Task3_01_ReadXlsx {

    @Test
    public void readXlsx() throws Exception {
        // 需要解析的xlsx文件
        String fileName = FileUtils.getPath() + "product.xlsx";
        // 用来临时存储解压缩目录
        String tempOutFilePath = FileUtils.getPath() + "productTemp/";

        try {
            //解压缩文件
            FileUtils.unZip(fileName, tempOutFilePath);
            // 解析 xl/sharedStrings.xml
            List<String> sharedStringList = readSharedStrings(tempOutFilePath);

            // 解析xl/worksheets/sheet1.xml
            readSheet1(tempOutFilePath, sharedStringList);
        } finally {
            // 删除文件要写到 finally
            FileUtils.delete(tempOutFilePath);
        }
    }

    /**
     * 将 sharedStringList解析成一个list
     *
     * @param tempOutFilePath
     * @return
     */
    private List<String> readSharedStrings(String tempOutFilePath) throws Exception {
        List<String> sharedStringList = new ArrayList<>();
        // 解析 xl/sharedStrings.xml
        // sst 根目录
        // sst -> si 代表一条文件数据
        // sst -> si -> t 代表存储的文本信息
        String sharedStringsFile = tempOutFilePath + "xl/sharedStrings.xml";
        SAXReader reader = new SAXReader();
        File file = new File(sharedStringsFile);
        Document document = reader.read(file);
        // sst
        Element sst = document.getRootElement();
        // sst -> si
        List<Element> siList = sst.elements("si");
        for (Element si : siList) {
            // sst -> si -> t
            sharedStringList.add(si.element("t").getText());
        }
        return sharedStringList;
    }

    /**
     * 将shee1.xml 解析成一行一行的单元格数据 并输出
     *
     * @param tempOutFilePath
     * @param sharedStringList
     * @throws Exception
     */
    private void readSheet1(String tempOutFilePath, List<String> sharedStringList) throws Exception {
        // 解析xl/worksheets/sheet1.xml
        // worksheet -> sheetData 存储了所有数据信息
        // worksheet -> sheetData -> row 代表一行数据
        // worksheet -> sheetData -> row -> c 代表个单元格，里面的 r 标签代表所在列 ，标签 t="s"
        // 代表当前单元并没有存储数据，真正数据存储在xl/sharedStrings.xml里面
        // worksheet -> sheetData -> row -> c —> v 存储具体数据
        String sheet1File = tempOutFilePath + "xl/worksheets/sheet1.xml";
        SAXReader reader = new SAXReader();
        File file = new File(sheet1File);
        Document document = reader.read(file);
        // worksheet
        Element worksheet = document.getRootElement();
        // worksheet -> sheetData
        Element sheetData = worksheet.element("sheetData");
        // worksheet -> sheetData -> row
        List<Element> rowList = sheetData.elements("row");
        for (int i = 0; i < rowList.size(); i++) {
            Element row = rowList.get(i);
            //worksheet -> sheetData -> row -> c
            List<Element> cList = row.elements("c");
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.append("第").append(i).append("行数据为：");
            for (Element c : cList) {
                //  单元格存储的值 这里有2个情况： 1. 直接就是数据 2. 只是一个坐标具体数据需要去sharedStringList读取
                String value = c.element("v").getText();
                // c这个节点 上面 t的属性值
                String tAttributeValue = c.attributeValue("t");
                if ("s".equals(tAttributeValue)) {
                    stringBuilder.append(sharedStringList.get(Integer.parseInt(value)))
                        .append(",");
                } else {
                    stringBuilder.append(value)
                        .append(",");
                }

            }
            log.info(stringBuilder.toString());
        }
    }
}
