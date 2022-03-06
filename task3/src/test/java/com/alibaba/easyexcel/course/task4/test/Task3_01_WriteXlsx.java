package com.alibaba.easyexcel.course.task4.test;

import java.text.SimpleDateFormat;
import java.util.Date;

import com.alibaba.easyexcel.course.base.utils.DateUtils;
import com.alibaba.easyexcel.course.base.utils.FileUtils;
import com.alibaba.easyexcel.course.base.utils.PositionUtils;

import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

/**
 * 手动生成一个xlsx文件
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class Task3_01_WriteXlsx {

    /**
     * 写一个xlsx的文件
     * 任务说明：前面我们说过 xlsx实际上是一堆xml文件打成的一个zip包，所以这里需要把所有的xml都准备完成，然后打包。
     * 由于很多xml不需要我们记，所以我这里以writeBase方法来准备好了，包括最后压缩成的zip包，我这里也直接写完了。
     * 需要完成：
     * 1. 完成writeSheet1方法
     * 2. 写入头部数据
     * 3. 写入内容数据
     *
     * @throws Exception
     */
    @Test
    public void writeXlsx() throws Exception {
        // 需要生成的xls文件
        String fileName = FileUtils.getPath() + "demo" + System.currentTimeMillis() + ".xlsx";
        // 用来临时生成xml的文件
        String tempOutFilePath = FileUtils.getPath() + "write/demo" + System.currentTimeMillis() + "/";
        try {
            // 基础信息写入 excel 真正要运行会有很多基础文件 这里我们不一一解析 代码已经写好
            // 很我们还是关注 xl/worksheets/sheet1.xml
            writeBase(tempOutFilePath);

            // 往我们第一个sheet里面写入数据
            // TODO 这里需要我们去实现
            writeSheet1(tempOutFilePath);

            // 将临时文件夹压缩成xlsx文件
            FileUtils.zip(fileName, tempOutFilePath);
        } finally {
            // 删除文件要写到 finally
            FileUtils.delete(tempOutFilePath);
        }
        log.info("生成xlsx完成：{}", fileName);
    }

    /**
     * 将 sharedStringList解析成一个list
     *
     * @param tempOutFilePath
     * @return
     */
    private void writeSheet1(String tempOutFilePath) throws Exception {
        //  xl/worksheets/sheet1.xml
        //  写入表头信息
        FileUtils.writeStringToFile(tempOutFilePath + "xl/worksheets/sheet1.xml", writer -> {
            writer.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats"
                + ".org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats"
                + ".org/officeDocument/2006/relationships\"><sheetPr filterMode=\"false\"><pageSetUpPr "
                + "fitToPage=\"false\" autoPageBreaks=\"false\"/></sheetPr><dimension "
                + "ref=\"A1\"/><sheetViews><sheetView "
                + "workbookViewId=\"0\"></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15"
                + ".0\"/>");
            writer.append("<sheetData>");

            // 写入头数据
            // TODO 写入标题头 数据
            // 使用  writer.append拼接 这么一个字符串 <row r="1"><c r="A1" t="str"><v>string</v></c><c r="B1"
            // t="str"><v>date</v></c><c
            // r="C1" t="str"><v>integer</v></c></row>
            // 其中 row标签中的 r="1"  1代表排序 第几行
            // c 标签中的 r="A1" A 代表第几列 excel用A 表示第一列 B 表示第二列 以此类推  后面的1 表示 当前第几行 ,这里的转换有兴趣的同学
            // 可以自己实现一个，入参是行号和列号，出参是 A1之类的拼接完成的，当然也可以直接使用：PositionUtils.position
            // c 标签中的 t="str" 代表 这是一个字符串
            // 这里觉得模模糊糊的同学可以解压缩 demo.xlsx ，然后打开sheet1.xml 并 格式化后看看看是怎么样子的
            // 到这里完成后建议先运行下代码，看看生成的是否可以用office软件打开，如果不可以打开，或者提示需要修复之类的 需要解压后打开：sheet1.xml  看下生成的是否与demo里面的一致

            // TODO 写入内容数据
            // 使用  writer.append 拼接 第一行数据：<row r="2"><c r="A2" t="str"><v>标题1</v></c><c r="B2" s="1"><v>44562
            // .0</v></c><c r="C2" ><v>1</v></c></row>
            // 和标题中已经一致的数据不再重复讲述，重点关注日期和数字的那一列
            // 日期： c 标签 需要传入 s="1" 实际上就是取style的index=1 的格式 ，然后日期转换成v标签的值已经写好方法：DateUtils.convertToExcelDate
            // 数字： c 标签 不要额外的 参数
            // 接下来依次类推 拼接 第二行 第三行的数据，当然 我们这里最好要使用for循环去处理

            writer.append(buildRow(0, "string", "date", "integer"));

            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
            // 写入1-10行数据
            for (int i = 1; i <= 10; i++) {
                writer.append(buildRow(i, "标题" + i, simpleDateFormat.parse("2022-01-" + i), i));
            }

            // 写入尾部信息
            writer.append("</sheetData>");
            writer.append(
                "<phoneticPr fontId=\"1\" type=\"noConversion\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\""
                    + " bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" "
                    + "orientation=\"portrait\" horizontalDpi=\"0\" verticalDpi=\"0\"/></worksheet>");
        });
    }

    private String buildRow(int rowIndex, Object... objs) {
        StringBuilder row = new StringBuilder();
        row.append("<row r=\"").append(rowIndex + 1).append("\">");

        // 列号
        int columnIndex = 0;
        for (Object obj : objs) {
            row.append(buildCell(rowIndex, columnIndex++, obj));
        }

        row.append("</row>");
        return row.toString();
    }

    private String buildCell(int rowIndex, int column, Object data) {
        StringBuilder cell = new StringBuilder();
        cell.append("<c r=\"").append(PositionUtils.position(rowIndex, column)).append("\" ");

        if (data == null) {
            cell.append("></c>");
            return cell.toString();
        }
        Class<?> clazz = data.getClass();
        if (clazz == String.class) {
            // string 2种情况
            // 1. <c r="A1" t="s"> 这种情况下 v标签只放索引， 具体值在sharedStrings.xml
            // 2. <c r="A1" t="str"> 这中情况下值 直接放在v标签下面
            // 为了简单 我们直接用第二种
            cell.append("t=\"str\"><v>");
            cell.append(data);
        } else if (clazz == Date.class) {
            // 日期  <c r="A2" s="1">
            // s="1" 代表设置为1号样式 在style.xml 写了1是日期格式
            cell.append("s=\"1\"><v>");
            cell.append(DateUtils.convertToExcelDate((Date)data));
        } else if (clazz == Integer.class) {
            // 数字 <c r="A3">
            cell.append("><v>");
            cell.append(data);
        } else {
            throw new IllegalArgumentException("当前还不支持字段类型" + clazz);
        }
        cell.append("</v></c>");
        return cell.toString();
    }

    private void writeBase(String tempOutFilePath) throws Exception {
        //  [Content_Types].xml
        FileUtils.writeStringToFile(tempOutFilePath + "[Content_Types].xml", writer -> {
            writer.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas"
                + ".openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" "
                + "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default "
                + "Extension=\"xml\" ContentType=\"application/xml\"/>");
            writer.append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd"
                + ".openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/><Override "
                + "PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument"
                + ".spreadsheetml.styles+xml\"/><Override PartName=\"/xl/workbook.xml\" "
                + "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            writer.append("<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd"
                + ".openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");

            writer.append(
                "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package"
                    + ".core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd"
                    + ".openxmlformats-officedocument.extended-properties+xml\"/></Types>");
        });

        // docProps/app.xml
        FileUtils.writeStringToFile(tempOutFilePath + "docProps/app.xml", writer -> writer.append(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Properties xmlns=\"http://schemas.openxmlformats"
                + ".org/officeDocument/2006/extended-properties\"><Application>Easy Excel</Application><AppVersion>1"
                + ".0.0</AppVersion></Properties>"));

        // docProps/core.xml
        FileUtils.writeStringToFile(tempOutFilePath + "docProps/core.xml", writer -> writer.append(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><cp:coreProperties "
                + "xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
                + "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" "
                + "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dcterms:created "
                + "xsi:type=\"dcterms:W3CDTF\">2022-01-01T00:00:00"
                + ".000Z</dcterms:created><dc:creator>Easy Excel</dc:creator></cp:coreProperties>"));

        // _rels/.rels"
        FileUtils.writeStringToFile(tempOutFilePath + "_rels/.rels", writer -> writer.append(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas"
                + ".openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas"
                + ".openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app"
                + ".xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats"
                + ".org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core"
                + ".xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats"
                + ".org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook"
                + ".xml\"/></Relationships>"));

        // xl/workbook.xml
        FileUtils.writeStringToFile(tempOutFilePath + "xl/workbook.xml", writer -> writer.append(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><workbookPr "
                + "date1904=\"false\"/><bookViews><workbookView /></bookViews><sheets><sheet name=\"Sheet1\" "
                + "r:id=\"rId3\" sheetId=\"1\"/></sheets></workbook>"));

        // xl/_rels/workbook.xml.rels
        FileUtils.writeStringToFile(tempOutFilePath + "xl/_rels/workbook.xml.rels", writer -> {
            writer.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><Relationships "
                + "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" "
                + "Target=\"sharedStrings.xml\" Type=\"http://schemas.openxmlformats"
                + ".org/officeDocument/2006/relationships/sharedStrings\"/><Relationship Id=\"rId2\" Target=\"styles"
                + ".xml\" Type=\"http://schemas.openxmlformats"
                + ".org/officeDocument/2006/relationships/styles\"/><Relationship Id=\"rId3\" "
                + "Target=\"worksheets/sheet1.xml\" Type=\"http://schemas.openxmlformats"
                + ".org/officeDocument/2006/relationships/worksheet\"/></Relationships>");
        });

        // xl/styles.xml
        FileUtils.writeStringToFile(tempOutFilePath + "xl/styles.xml", writer -> {
            writer.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet xmlns=\"http://schemas"
                + ".openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"0\"></numFmts><fonts "
                + "count=\"1\"><font><sz val=\"11.00\"/><color rgb=\"FF000000\"/><name "
                + "val=\"Calibri\"/></font></fonts><fills count=\"2\"><fill><patternFill "
                + "patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders "
                + "count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs "
                + "count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs "
                + "count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"><alignment "
                + "vertical=\"center\"/></xf><xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" "
                + "applyNumberFormat=\"1\"><alignment vertical=\"center\"/></xf></cellXfs><dxfs "
                + "count=\"0\"></dxfs></styleSheet>");
        });

        // xl/sharedStrings.xml
        FileUtils.writeStringToFile(tempOutFilePath + "xl/sharedStrings.xml", writer -> {
            writer.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"http://schemas"
                + ".openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"></sst>");
        });
    }

}
