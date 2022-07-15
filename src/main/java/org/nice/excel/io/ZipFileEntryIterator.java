package org.nice.excel.io;

import lombok.SneakyThrows;
import org.apache.commons.compress.archivers.zip.ZipArchiveInputStream;
import org.nice.excel.ExcelException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import static org.nice.excel.io.FileIteratorConstant.EXCEL_CHARSET;


/**
 * <p>
 * zip格式的{@link FileEntryIterator}，用于从zip文件中迭代{@link FileEntry}
 * </p>
 *
 * @author suzh@tsintergy.com
 * @date 2020/10/14 14:28
 */
public class ZipFileEntryIterator implements FileEntryIterator {

    private static final Logger LOGGER = LoggerFactory.getLogger(ZipFileEntryIterator.class);

    /**
     * zip格式的输入流，在读取数据时只会读取当前ZipEntry对应的数据，要想读取下一个ZipEntry中的数据需要调用{@link ZipInputStream#closeEntry()}
     */
    private final ZipArchiveInputStream zipInputStream;

    /**
     * 当前迭代到的ZipEntry
     */
    private ZipEntry currentZipEntry;

    public ZipFileEntryIterator(File file) throws FileNotFoundException {
        this(new FileInputStream(file));
    }

    public ZipFileEntryIterator(InputStream inputStream) {
//        使用ZipArchiveInputStream处理：java.util.zip.ZipException: only DEFLATED entries can have EXT descriptor
//        https://stackoverflow.com/questions/15738312/how-to-fix-org-apache-commons-compress-archivers-zip-unsupportedzipfeatureexcept
        zipInputStream = new ZipArchiveInputStream(inputStream, EXCEL_CHARSET, true, true);
    }

    @SneakyThrows
    @Override
    public boolean hasNext() {
        // 循环找到下一个ZipEntry，当ZipEntry不存在或者非目录时，标记为currentZipEntry，跳出循环
        try {
            while (true) {
                ZipEntry entry = zipInputStream.getNextZipEntry();
                if (entry == null || !entry.isDirectory()) {
                    currentZipEntry = entry;
                    break;
                }
            }
        } catch (IOException e) {
            throw new ExcelException(e);
        }

        // 如果经过前面的操作，currentZipEntry不存在，则说明已经读取到流的末尾了，关闭流
        if (currentZipEntry == null) {
            try {
                zipInputStream.close();
            } catch (IOException e) {
                LOGGER.error("关闭ZipInputStream出现异常!");
            }
        }

        return currentZipEntry != null;
    }

    @Override
    public FileEntry next() {
        // 将currentZipEntry的名字和对应的流包装成ExcelEntry
        return new FileEntry(currentZipEntry.getName(), zipInputStream);
    }
}
