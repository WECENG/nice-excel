package org.nice.excel.io;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;

/**
 * <p>
 * 未压缩的文件迭代器
 * </p>
 *
 * @author WECENG
 * @since 2021/3/2 11:13
 */
public class RawFileEntryIterator implements FileEntryIterator {

    private static final Logger LOGGER = LoggerFactory.getLogger(RawFileEntryIterator.class);

    private boolean close;

    private final FileEntry fileEntry;

    public RawFileEntryIterator(FileEntry fileEntry) {
        this.fileEntry = fileEntry;
    }

    @Override
    public boolean hasNext() {
        if (getClose()) {
            try {
                fileEntry.getExcelInputStream().close();
            } catch (IOException e) {
                LOGGER.error("IO流关闭异常!");
            }
        }
        return !getClose();
    }

    @Override
    public FileEntry next() {
        setClose(true);
        return fileEntry;
    }

    public boolean getClose() {
        return close;
    }

    public void setClose(boolean close) {
        this.close = close;
    }
}
