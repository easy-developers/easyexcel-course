package com.alibaba.easyexcel.course.task3;

import java.io.BufferedWriter;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

import com.alibaba.easyexcel.course.base.utils.DateUtils;
import com.alibaba.easyexcel.course.base.utils.FileUtils;
import com.alibaba.easyexcel.course.base.utils.PositionUtils;

import net.sf.cglib.beans.BeanMap;

public class MyExcelWriterImpl<T> implements MyExcelWriter<T> {

    /**
     * 用来存储临时数据
     */
    private static final String MY_EXCEL_TEMP_DIR = System.getProperty("java.io.tmpdir") + "MyExcelWrite/";

    /**
     * 输出的文件地址
     */
    private String file;
    /**
     * clazz的字段
     */
    private List<Field> fieldList;

    /**
     * 生成临时文件的地址
     */
    private String tempOutFilePath;
    /**
     * 用于写sheet1文件
     */
    private BufferedWriter sheet1Writer;

    /**
     * 行号
     */
    private int rowIndex;

    public MyExcelWriterImpl(String file, Class<T> clazz) throws Exception {
        this.file = file;
        this.fieldList = Arrays.asList(clazz.getDeclaredFields());
        rowIndex = 0;
        // 用来临时生成xml的文件
        tempOutFilePath = MY_EXCEL_TEMP_DIR + System.currentTimeMillis() + "/";
        // 基础信息写入 excel 真正要运行会有很多基础文件 这里我们不一一解析 代码已经写好
        // 很我们还是关注 xl/worksheets/sheet1.xml
        writeBase();

        // 写入在写入行数据之前的数据
        writeSheet1BeforeRowData();
    }

    /**
     * 写入在写入行数据之前的数据
     *
     * @return
     */
    private void writeSheet1BeforeRowData() throws Exception {
        //  xl/worksheets/sheet1.xml
        sheet1Writer = FileUtils.buildWriter(tempOutFilePath + "xl/worksheets/sheet1.xml");

        sheet1Writer.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats"
            + ".org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats"
            + ".org/officeDocument/2006/relationships\"><sheetPr filterMode=\"false\"><pageSetUpPr "
            + "fitToPage=\"false\" autoPageBreaks=\"false\"/></sheetPr><dimension "
            + "ref=\"A1\"/><sheetViews><sheetView "
            + "workbookViewId=\"0\"></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15"
            + ".0\"/>");
        sheet1Writer.append("<sheetData>");

        // 行号
        // 写入头数据
        sheet1Writer.append(buildRow(fieldList.stream()
            .map(Field::getName)
            .collect(Collectors.toList())));
    }

    /**
     * 写入在写入行数据之后的数据
     *
     * @return
     */
    private void afterSheet1BeforeRowData() throws Exception {
        //  xl/worksheets/sheet1.xml
        sheet1Writer.append("</sheetData>");
        sheet1Writer.append(
            "<phoneticPr fontId=\"1\" type=\"noConversion\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\""
                + " bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" "
                + "orientation=\"portrait\" horizontalDpi=\"0\" verticalDpi=\"0\"/></worksheet>");
        sheet1Writer.close();
    }

    private String buildRow(List<?> objs) {
        StringBuilder row = new StringBuilder();
        row.append("<row r=\"").append(rowIndex + 1).append("\">");

        // 列号
        int columnIndex = 0;
        for (Object obj : objs) {
            row.append(buildCell(columnIndex++, obj));
        }

        row.append("</row>");
        rowIndex++;
        return row.toString();
    }

    private String buildCell(int column, Object data) {
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

    private void writeBase() throws Exception {
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

    @Override
    public void write(List<T> dataList) throws Exception {
        if (dataList == null || dataList.isEmpty()) {
            return;
        }
        for (T data : dataList) {
            // 创建一个cglib对象
            BeanMap beanMap = BeanMap.create(data);
            // 将数据放入对象
            sheet1Writer.append(buildRow(fieldList.stream()
                .map(field -> beanMap.get(field.getName()))
                .collect(Collectors.toList())));
        }
    }

    @Override
    public void finish() throws Exception {
        try {
            //  完成sheet.xml
            afterSheet1BeforeRowData();

            // 将临时文件夹压缩成xlsx文件
            FileUtils.zip(file, tempOutFilePath);
        } finally {
            // 删除临时文件地址
            FileUtils.delete(tempOutFilePath);
        }

    }

    @Override
    public void close() throws Exception {
        finish();
    }
}
