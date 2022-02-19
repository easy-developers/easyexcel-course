package com.alibaba.easyexcel.course.task3.test;

import com.alibaba.easyexcel.course.base.utils.FileUtils;

import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

/**
 * 任务目标： 读取中xlsx中的信息
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class Task4_01_WriteXlsx {

    @Test
    public void writeXlsx() throws Exception {
        // 需要生成的xls文件
        String fileName = FileUtils.getPath() + "product" + System.currentTimeMillis() + ".xlsx";
        // 用来临时生成xml的文件
        String tempOutFilePath = FileUtils.getPath() + "write/excelTemp/";
        // 基础信息写入
        //writeBase();

    }

    private void writeBase(String tempOutFilePath) {
        ////  [Content_Types].xml
        //FileUtils.writeStringToFile(tempOutFilePath + "[Content_Types].xml", writer -> {
        //    writer.append(
        //        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas"
        //            + ".openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" "
        //            + "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default "
        //            + "Extension=\"xml\" ContentType=\"application/xml\"/>");
        //    w.append(
        //        "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd"
        //            + ".openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/><Override "
        //            + "PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument"
        //            + ".spreadsheetml.styles+xml\"/><Override PartName=\"/xl/workbook.xml\" "
        //            + "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
        //    for (Worksheet ws : worksheets) {
        //        int index = getIndex(ws);
        //        w.append("<Override PartName=\"/xl/worksheets/sheet").append(index).append(
        //            ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml"
        //                + ".worksheet+xml\"/>");
        //
        //    }
        //    w.append(
        //        "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package"
        //            + ".core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>");
        //});
        //
        //FileUtils.writeStringToFile();
    }

}
