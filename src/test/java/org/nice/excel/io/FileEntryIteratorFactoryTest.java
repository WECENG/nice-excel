package org.nice.excel.io;

import org.junit.jupiter.api.Test;
import org.nice.excel.ExcelException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;

import static org.junit.jupiter.api.Assertions.assertNotNull;


public class FileEntryIteratorFactoryTest {

    private final Logger log = LoggerFactory.getLogger(this.getClass());

    @Test
    public void createZip() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/all.zip");
        FileEntryIterator fileEntryIterator = FileEntryIteratorFactory.create("all.zip", in);
        while (fileEntryIterator.hasNext()) {
            FileEntry fileEntry = fileEntryIterator.next();
            assertNotNull(fileEntry);
            log.info("文件名称：{}", fileEntry.getName());
        }
    }

    @Test
    public void createRar() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/all.rar");
        FileEntryIterator fileEntryIterator = FileEntryIteratorFactory.create("all.rar", in);
        while (fileEntryIterator.hasNext()) {
            FileEntry fileEntry = fileEntryIterator.next();
            assertNotNull(fileEntry);
            log.info("文件名称：{}", fileEntry.getName());
        }
    }
}