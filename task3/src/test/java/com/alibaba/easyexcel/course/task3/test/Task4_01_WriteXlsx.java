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
        String tempOutFilePath = FileUtils.getPath() + "write/excelTemp" + System.currentTimeMillis() + "/";
        try {
            // 基础信息写入 excel 真正要运行会有很多基础文件 这里我们不一一解析 代码已经写好
            // 很我们还是关注 xl/worksheets/sheet1.xml
            writeBase(tempOutFilePath);

            writeSheet1(tempOutFilePath);

            FileUtils.zip(fileName, tempOutFilePath);
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
    private void writeSheet1(String tempOutFilePath) throws Exception {
        //  xl/worksheets/sheet1.xml
        FileUtils.writeStringToFile(tempOutFilePath + "xl/worksheets/sheet1.xml", writer -> {
            writer.append(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats"
                    + ".org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats"
                    + ".org/officeDocument/2006/relationships\"><sheetPr filterMode=\"false\"><pageSetUpPr "
                    + "fitToPage=\"false\" autoPageBreaks=\"false\"/></sheetPr><dimension "
                    + "ref=\"A1\"/><sheetViews><sheetView "
                    + "workbookViewId=\"0\"></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15"
                    + ".0\"/><cols><col min=\"1\" max=\"1\" width=\"3.42578125\" customWidth=\"true\" "
                    + "bestFit=\"true\"/></cols>");
            writer.append("<sheetData>");

            writer.append("<row r=\"1\"><c r=\"A1\" t=\"str\"><v>张三</v></c></row>");

            writer.append("</sheetData>");
            writer.append(
                "<phoneticPr fontId=\"1\" type=\"noConversion\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\""
                    + " bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" "
                    + "orientation=\"portrait\" horizontalDpi=\"0\" verticalDpi=\"0\"/></worksheet>");
        });

    }

    private void writeBase(String tempOutFilePath) throws Exception {
        //  [Content_Types].xml
        FileUtils.writeStringToFile(tempOutFilePath + "[Content_Types].xml", writer -> {
            writer.append(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas"
                    + ".openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" "
                    + "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default "
                    + "Extension=\"xml\" ContentType=\"application/xml\"/>");
            writer.append(
                "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd"
                    + ".openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/><Override "
                    + "PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument"
                    + ".spreadsheetml.styles+xml\"/><Override PartName=\"/xl/workbook.xml\" "
                    + "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            writer.append(
                "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd"
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
        FileUtils.writeStringToFile(tempOutFilePath + "xl/workbook.xml",
            writer -> writer.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
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
                + "count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" "
                + "xfId=\"0\"/></cellXfs><dxfs count=\"0\"></dxfs></styleSheet>");
        });

        // xl/sharedStrings.xml
        FileUtils.writeStringToFile(tempOutFilePath + "xl/sharedStrings.xml", writer -> {
            writer.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"http://schemas"
                + ".openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"></sst>");
        });
    }

}
