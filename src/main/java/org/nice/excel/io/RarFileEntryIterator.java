package org.nice.excel.io;

import com.github.junrar.Archive;
import com.github.junrar.exception.UnsupportedRarV5Exception;
import com.github.junrar.rarfile.FileHeader;
import lombok.SneakyThrows;
import org.nice.excel.ExcelException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.Objects;

/**
 * <p>
 * rar格式的{@link FileEntryIterator}，用于从rar文件中迭代{@link FileEntry}
 * </p>
 *
 * @author WECENG
 * @since 2021/3/2 10:07
 */
public class RarFileEntryIterator implements FileEntryIterator {

    private static final Logger LOGGER = LoggerFactory.getLogger(RarFileEntryIterator.class);

    private final Archive archive;

    private FileHeader currentFileHeader;

    public RarFileEntryIterator(InputStream inputStream) throws ExcelException {
        try {
            archive = new Archive(inputStream);
        } catch (UnsupportedRarV5Exception rarV5Exception) {
            throw new ExcelException("暂不支持该版本的压缩文件！");
        } catch (Exception e) {
            throw new ExcelException(e);
        }
    }

    public RarFileEntryIterator(File file) throws FileNotFoundException, ExcelException {
        this(new FileInputStream(file));
    }

    @Override
    public boolean hasNext() {
        while (true) {
            FileHeader nextFileHeader = archive.nextFileHeader();
            if (Objects.isNull(nextFileHeader) || !nextFileHeader.isDirectory()) {
                currentFileHeader = nextFileHeader;
                break;
            }
        }
        //close
        if (Objects.isNull(currentFileHeader)) {
            try {
                archive.close();
            } catch (IOException e) {
                LOGGER.error("IO流关闭异常!");
            }
        }
        return Objects.nonNull(currentFileHeader);
    }

    @SneakyThrows
    @Override
    public FileEntry next() {
        try {
            return new FileEntry(currentFileHeader.getFileName(), archive.getInputStream(currentFileHeader));
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }
}
