package com.alibaba.easyexcel.course.base.utils.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import com.alibaba.easyexcel.course.base.utils.FileUtils;

import com.github.rzymek.opczip.OpcZipOutputStream;
import com.github.rzymek.opczip.reader.skipping.ZipStreamReader;
import org.junit.Test;

public class FileUtilsTest {

    @Test
    public void zip() throws Exception {
        String pathdd = "/Users/zhuangjiaju/Downloads/tet3/product07";

        try (ZipOutputStream out = new ZipOutputStream(
            new FileOutputStream("/Users/zhuangjiaju/Downloads/tet3/ddd3.xlsx"))) {
            File[] flist = new File(pathdd).listFiles();
            for (File file : flist) {
                if (file.isDirectory()) {
                    continue;
                }
                out.putNextEntry(new ZipEntry(file.getName()));
                Files.copy(Paths.get(file.getAbsolutePath()), out);
                out.closeEntry();
            }
        }
    }

    @Test
    public void zip2() throws Exception {
        String pathdd = "/Users/zhuangjiaju/Downloads/tet3/product07";
        FileUtils.zip("/Users/zhuangjiaju/Downloads/tet3/dd" + System.currentTimeMillis() + ".xlsx", pathdd);
    }

    @Test
    public void zip3() throws Exception {
        String pathdd = "/Users/zhuangjiaju/Downloads/tet3/excelTemp1645341070311";
        List<String> paths = Arrays.asList(
            "[Content_Types].xml",
            "_rels/.rels",
            "docProps/app.xml",
            "docProps/core.xml",
            "xl/styles.xml",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/worksheets/sheet1.xml"
        );
        try (OpcZipOutputStream out = new OpcZipOutputStream(
            new FileOutputStream("/Users/zhuangjiaju/Downloads/tet3/rrr" + System.currentTimeMillis() + ".xlsx"))) {
            for (String path : paths) {
                out.putNextEntry(new ZipEntry(path));
                Files.copy(Paths.get(pathdd, path), out);
                out.closeEntry();
            }
        }
    }

    @Test
    public void zip31() throws Exception {
        String pathdd = "/Users/zhuangjiaju/Downloads/tet3/right";
        List<String> paths = Arrays.asList(
            "[Content_Types].xml",
            "_rels/.rels",
            "docProps/app.xml",
            "docProps/core.xml",
            "xl/styles.xml",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/worksheets/sheet1.xml"
        );
        try (OpcZipOutputStream out = new OpcZipOutputStream(
            new FileOutputStream("/Users/zhuangjiaju/Downloads/tet3/eee" + System.currentTimeMillis() + ".xlsx"))) {
            for (String path : paths) {
                out.putNextEntry(new ZipEntry(path));
                Files.copy(Paths.get(pathdd, path), out);
                out.closeEntry();
            }
        }
    }

    @Test
    public void zip4() throws Exception {
        String file = "/Users/zhuangjiaju/Downloads/tet3/product07.xlsx";

        String pathdd = "/Users/zhuangjiaju/Downloads/tet3/product07" + System.currentTimeMillis();

        try (ZipStreamReader reader = new ZipStreamReader(new FileInputStream(file))) {
            byte[] sharedZipped = null;
            for (; ; ) {
                ZipEntry entry = reader.nextEntry();
                if (entry == null) {
                    break;
                }
                if (entry.isDirectory()) {
                    String dirPath = pathdd + "/" + entry.getName();
                    new File(dirPath).mkdirs();
                } else {
                    File targetFile = new File(pathdd + "/" + entry.getName());
                    if (!targetFile.getParentFile().exists()) {
                        targetFile.getParentFile().mkdirs();
                    }
                    InputStream is = reader.getUncompressedStream();
                    try (FileOutputStream fos = new FileOutputStream(targetFile)) {
                        int len;
                        byte[] buf = new byte[1024];
                        while ((len = is.read(buf)) != -1) {
                            fos.write(buf, 0, len);
                        }
                    }
                }
            }
        }

        List<String> paths = Arrays.asList(
            "[Content_Types].xml",
            "_rels/.rels",
            "docProps/app.xml",
            "docProps/core.xml",
            "xl/styles.xml",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/worksheets/sheet1.xml"
        );
        try (OpcZipOutputStream out = new OpcZipOutputStream(
            new FileOutputStream("/Users/zhuangjiaju/Downloads/tet3/eee" + System.currentTimeMillis() + ".xlsx"))) {
            for (String path : paths) {
                out.putNextEntry(new ZipEntry(path));
                Files.copy(Paths.get(pathdd, path), out);
                out.closeEntry();
            }
        }
    }
}
