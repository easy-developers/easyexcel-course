package com.alibaba.easyexcel.course.base.utils;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import ch.qos.logback.core.util.FileUtil;
import com.github.rzymek.opczip.OpcZipOutputStream;
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
     * @param outputPath  输出路径
     */
    public static void unZip(String zipFileName, String outputPath) throws IOException {
        ZipFile zipFile = new ZipFile(zipFileName);
        Enumeration<?> entries = zipFile.entries();
        while (entries.hasMoreElements()) {
            ZipEntry entry = (ZipEntry)entries.nextElement();
            if (entry.isDirectory()) {
                String dirPath = outputPath + "/" + entry.getName();
                new File(dirPath).mkdirs();
            } else {
                File targetFile = new File(outputPath + "/" + entry.getName());
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
     * 将数据写入到一个文件
     *
     * @param fileName 需要写如到的文件¬
     * @param writer
     */
    public static void writeStringToFile(String fileName, ConsumerThrowsException<BufferedWriter> writer)
        throws Exception {
        File file = new File(fileName);
        if (!file.exists()) {
            file.createNewFile();
        }
        try (BufferedWriter out = new BufferedWriter(new FileWriter(file))) {
            writer.accept(out);
        }
    }

    /**
     * 压缩一个文件
     *
     * @param zipFileName 需要压缩的文件
     * @param inputPath   需要压缩的路径
     */
    public static void zip(String zipFileName, String inputPath) throws IOException {
        try (OpcZipOutputStream out = new OpcZipOutputStream(new FileOutputStream(zipFileName))) {
            File input = new File(inputPath);
            compress(out, input, null);
            out.finish();
        }
    }

    /**
     * 压缩文件
     *
     * @param out
     * @param input
     * @param name
     * @throws IOException
     */
    private static void compress(OpcZipOutputStream out, File input, String name)
        throws IOException {
        if (name == null) {
            name = input.getName();
        }
        if (input.isDirectory()) {
            File[] flist = input.listFiles();
            if (flist == null || flist.length == 0) {
                out.putNextEntry(new ZipEntry(name + "/"));
                out.closeEntry();
            } else {
                for (File file : flist) {
                    compress(out, file, name + "/" + file.getName());
                }
            }
        } else {
            out.putNextEntry(new ZipEntry(name));
            Files.copy(Paths.get(input.getAbsolutePath()), out);
            out.closeEntry();
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
