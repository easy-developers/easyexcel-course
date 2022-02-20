package com.alibaba.easyexcel.course.task4;

import java.io.File;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.Collectors;

import com.alibaba.easyexcel.course.base.utils.DateUtils;
import com.alibaba.easyexcel.course.base.utils.FileUtils;

import lombok.extern.slf4j.Slf4j;
import net.sf.cglib.beans.BeanMap;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

@Slf4j
public class MyExcelReaderImpl<T> implements MyExcelReader<T> {

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
        SAXReader reader = new SAXReader();
        File file = new File(sheet1File);
        Document document = reader.read(file);
        // worksheet
        Element worksheet = document.getRootElement();
        // worksheet -> sheetData
        Element sheetData = worksheet.element("sheetData");
        // worksheet -> sheetData -> row
        List<Element> rowList = sheetData.elements("row");
        // 存储所有数据
        List<T> dataList = new ArrayList<>();
        for (int i = 0; i < rowList.size(); i++) {
            if (i == 0) {
                log.info("第一条为数据头，忽略");
                continue;
            }
            Element row = rowList.get(i);
            //worksheet -> sheetData -> row -> c
            List<Element> cList = row.elements("c");

            // 创建一个实例对象
            T data = clazz.newInstance();
            dataList.add(data);
            BeanMap beanMap = BeanMap.create(data);
            int column = 0;
            for (Element c : cList) {
                // 这列数据 已经超过我们对象了 后面的数据 直接忽略
                if (column >= fieldList.size()) {
                    break;
                }
                // 拿到当前列的 对应实体对象的属性
                Field field = fieldList.get(column);

                //  单元格存储的值 这里有2个情况： 1. 直接就是数据 2. 只是一个坐标具体数据需要去sharedStringList读取
                String value = c.element("v").getText();
                // c这个节点 上面 t的属性值
                String tAttributeValue = c.attributeValue("t");

                String dataValue;
                if ("s".equals(tAttributeValue)) {
                    dataValue = sharedStringList.get(Integer.parseInt(value));
                } else {
                    dataValue = value;
                }

                // 转换数据
                Object convertedValue = doConvert(dataValue, field);

                beanMap.put(field.getName(), convertedValue);
                column++;
            }
            // 每隔100条 回调异常 防止数据都在内存
            if (dataList.size() % 100 == 0) {
                consumer.accept(dataList);
                dataList = new ArrayList<>();
            }
        }

        // 可能最后还有一些数据回调
        consumer.accept(dataList);
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
}
