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

/**
 * 读取excel的工具
 *
 * @param <T>
 * @author Jiaju Zhuang
 */
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

            // TODO 解析 xl/sharedStrings.xml
            List<String> sharedStringList = readSharedStrings(tempOutFilePath);

            // TODO 解析xl/worksheets/sheet1.xml
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
        // 类似于上面的结构 要解析成这样子的一个数组 ["string","date"]
        // TODO 用 dom4j 解析 xl/sharedStrings.xml
        //不会使用dom4j的同学 可以 搜索： dom4j 解析xml

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
        // 代表当前单元并没有存储数据，真正数据存储在xl/sharedStrings.xml里面，只要将 v 标签中的素读取到，然后转成int ,然后 去sharedStringList 读取指定的index的数据即可
        // worksheet -> sheetData -> row -> c —> v 存储具体数据
        // 不会 同学可以打开： 解压缩的 xl/worksheets/sheet1.xml   并格式化一下的内容
        // TODO 用 dom4j 解析xl/worksheets/sheet1.xml 并封装成 T 对象后 回调给consumer.accept
        // 需要注意 在写入到对应的实体的时候： 需要根据不同的字段类型来转换 目前只要支持： string integer date即可 其中日期可以调用：DateUtils.convertToJavaDate
        // 将一个数字转成日期
        // 默认数据到达100条以后 ，回调：consumer.accept 给用户去处理数据，防止内存存储太多数据

        String sheet1File = tempOutFilePath + "xl/worksheets/sheet1.xml";

    }

}
