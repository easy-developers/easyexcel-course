package com.alibaba.easyexcel.course.task4;

import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.Collectors;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import com.alibaba.easyexcel.course.base.utils.DateUtils;
import com.alibaba.easyexcel.course.base.utils.FileUtils;

import lombok.extern.slf4j.Slf4j;
import net.sf.cglib.beans.BeanMap;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

/**
 * sax 读取
 *
 * @param <T>
 * @author Jiaju Zhuang
 */
@Slf4j
public class MyExcelReaderSaxImpl<T> implements MyExcelReader<T> {

    /**
     * 用来存储临时数据
     */
    private static final String MY_EXCEL_TEMP_DIR = System.getProperty("java.io.tmpdir") + "MyExcelRead/";

    @Override
    public void read(String file, Class<T> clazz, Consumer<List<T>> consumer) {
        // 生成一个临时文件的目录
        String tempOutFilePath = MY_EXCEL_TEMP_DIR + System.currentTimeMillis() + "/";

        try {
            //解压缩文件
            FileUtils.unZip(file, tempOutFilePath);

            // 解析 clazz 生成一个 List<字段的fieldName>
            // 这个用于 代表第几列 需要放到哪个属性里面
            // 这里如果需要性能更好 可以对这个list 做一个缓存 因为clazz 上线以后是不会变的
            List<Field> fieldList = Arrays.stream(clazz.getDeclaredFields())
                .collect(Collectors.toList());

            // 解析 xl/sharedStrings.xml
            // 这里数据量不会很大 所以 暂时不用sax读取了
            List<String> sharedStringList = readSharedStrings(tempOutFilePath);

            // 解析xl/worksheets/sheet1.xml
            readSheet1(tempOutFilePath, sharedStringList, clazz, fieldList, consumer);

        } catch (Exception e) {
            log.info("读取excel失败", e);
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
        String sharedStringsFile = tempOutFilePath + "xl/sharedStrings.xml";
        // sharedStrings 差不多是这样子的一个格式 <sst><si><t>string</t></si><si><t>date</t></si></sst>
        // sst 根目录
        // sst -> si 代表一条文件数据
        // sst -> si -> t 代表存储的文本信息
        // 也可以自己解压：demoWithSharedStrings.xlsx来看具体的信息
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
    private void readSheet1(String tempOutFilePath, List<String> sharedStringList, Class<T> clazz,
        List<Field> fieldList, Consumer<List<T>> consumer)
        throws Exception {
        // 解析xl/worksheets/sheet1.xml
        // worksheet -> sheetData 存储了所有数据信息
        // worksheet -> sheetData -> row 代表一行数据
        // worksheet -> sheetData -> row -> c 代表个单元格，里面的 r 标签代表所在列 ，标签 t="s"
        // 代表当前单元并没有存储数据，真正数据存储在xl/sharedStrings.xml里面
        // worksheet -> sheetData -> row -> c —> v 存储具体数据
        String sheet1File = tempOutFilePath + "xl/worksheets/sheet1.xml";

        try (FileInputStream fileInputStream = new FileInputStream(sheet1File)) {
            InputSource inputSource = new InputSource(fileInputStream);
            SAXParserFactory saxFactory = SAXParserFactory.newInstance();
            saxFactory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            saxFactory.setFeature("http://xml.org/sax/features/external-general-entities", false);
            saxFactory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
            SAXParser saxParser = saxFactory.newSAXParser();
            XMLReader xmlReader = saxParser.getXMLReader();
            xmlReader.setContentHandler(new MyDefaultHandler(clazz, fieldList, sharedStringList, consumer));
            xmlReader.parse(inputSource);
        }
    }

    private Object doConvert(String value, Field field) {
        if (value == null) {
            return null;
        }
        Class<?> fieldType = field.getType();
        if (fieldType == String.class) {
            return value;
        } else if (fieldType == Date.class) {
            return DateUtils.convertToJavaDate(value);
        } else if (fieldType == Integer.class) {
            return Integer.valueOf(value);
        } else {
            throw new IllegalArgumentException("当前还不支持字段类型" + field.getType());
        }
    }

    private class MyDefaultHandler extends DefaultHandler {

        private Class<T> clazz;
        private List<Field> fieldList;
        private List<String> sharedStringList;
        private Consumer<List<T>> consumer;

        private List<T> dataList;

        private BeanMap currentBeanMap;
        private int currentColumnIndex;
        private boolean currentSharedString;
        private StringBuilder currentString;
        private Field currentField;

        public MyDefaultHandler(Class<T> clazz, List<Field> fieldList, List<String> sharedStringList,
            Consumer<List<T>> consumer) {
            this.clazz = clazz;
            this.fieldList = fieldList;
            this.sharedStringList = sharedStringList;
            this.consumer = consumer;
            dataList = new ArrayList<>();
        }

        /**
         * 读到 任何一个 Element 的开始的时候 会回调这个方法
         *
         * @param uri
         * @param localName
         * @param name
         * @param attributes
         */
        @Override
        public void startElement(String uri, String localName, String name, Attributes attributes) {
            if ("row".equals(name)) {
                // 代表开启了一行
                try {
                    T data = clazz.newInstance();
                    dataList.add(data);
                    currentBeanMap = BeanMap.create(data);
                } catch (Exception e) {
                    log.error("创建实例异常");
                }
                currentColumnIndex = 0;
            } else if ("c".equals(name)) {
                // 这列数据 已经超过我们对象了 后面的数据 直接忽略
                if (currentColumnIndex >= fieldList.size()) {
                    return;
                }
                // 拿到当前列的 对应实体对象的属性
                currentField = fieldList.get(currentColumnIndex);
                // c这个节点 上面 t的属性值
                String tAttributeValue = attributes.getValue("t");
                if ("s".equals(tAttributeValue)) {
                    currentSharedString = true;
                } else {
                    currentSharedString = false;
                }
            } else if ("v".equals(name)) {
                currentString = new StringBuilder();
            }
        }

        /**
         * 读到 任何一个 Element 的介绍的时候 会回调这个方法
         *
         * @param uri
         * @param localName
         * @param name
         */
        @Override
        public void endElement(String uri, String localName, String name) {
            if ("row".equals(name)) {
                // 每隔100条 回调异常 防止数据都在内存
                if (dataList.size() % 100 == 0) {
                    consumer.accept(dataList);
                    dataList = new ArrayList<>();
                }
            } else if ("c".equals(name)) {
                // 转换数据
                String dataValue = currentString.toString();
                if (currentSharedString) {
                    dataValue = sharedStringList.get(Integer.parseInt(dataValue));
                }
                Object convertedValue = doConvert(dataValue, currentField);
                currentBeanMap.put(currentField.getName(), convertedValue);
                currentColumnIndex++;
            } else if ("sheetData".equals(name)) {
                // 结束当前表格读取
                consumer.accept(dataList);
            }
        }

        /**
         * Element有数据的时候 会回调这个方法
         *
         * @param ch
         * @param start
         * @param length
         */
        @Override
        public void characters(char[] ch, int start, int length) {
            currentString.append(ch, start, length);
        }
    }
}
