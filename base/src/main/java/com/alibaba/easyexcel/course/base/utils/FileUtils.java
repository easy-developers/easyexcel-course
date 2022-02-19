package com.alibaba.easyexcel.course.base.utils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import ch.qos.logback.core.util.FileUtil;
import lombok.extern.slf4j.Slf4j;

/**
 * 文件工具类
 *
 * @author Jiaju Zhuang
 */
@Slf4j
public class FileUtils {

    /**
     * 获取项目的的路径
     *
     * @return
     */
    public static String getPath() {
        return FileUtil.class.getResource("/").getPath();
    }

    /**
     * 解压缩一个文件
     *
     * @param zipFileName 需要解压缩的文件
     * @param outPutPath  输出路径
     */
    public static void unZip(String zipFileName, String outPutPath) throws IOException {
        ZipFile zipFile = new ZipFile(zipFileName);
        Enumeration<?> entries = zipFile.entries();
        while (entries.hasMoreElements()) {
            ZipEntry entry = (ZipEntry)entries.nextElement();
            if (entry.isDirectory()) {
                String dirPath = outPutPath + "/" + entry.getName();
                new File(dirPath).mkdirs();
            } else {
                File targetFile = new File(outPutPath + "/" + entry.getName());
                // 保证这个文件的父文件夹必须要存在
                if (!targetFile.getParentFile().exists()) {
                    targetFile.getParentFile().mkdirs();
                }
                targetFile.createNewFile();
                try (InputStream is = zipFile.getInputStream(entry);
                     FileOutputStream fos = new FileOutputStream(targetFile)) {
                    int len;
                    byte[] buf = new byte[1024];
                    while ((len = is.read(buf)) != -1) {
                        fos.write(buf, 0, len);
                    }
                }
            }
        }
    }

    /**
     * 删除一个文件夹以及所有文件
     *
     * @param fullFileOrDirPath
     * @return
     */
    public static boolean delete(String fullFileOrDirPath) {
        return delete(new File(fullFileOrDirPath));
    }

    /**
     * 删除一个文件夹或者文件以及所有文件
     *
     * @param file
     * @return
     */
    public static boolean delete(File file) {
        if (file == null || !file.exists()) {
            return true;
        }
        if (file.isDirectory()) {
            File[] files = file.listFiles();
            if (files != null) {
                for (File childFile : files) {
                    if (!delete(childFile)) {
                        return false;
                    }
                }
            }
        }
        return file.delete();
    }

}
