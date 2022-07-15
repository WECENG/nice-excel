package org.nice.excel.io;

import org.nice.excel.ExcelException;
import org.springframework.util.Assert;

import java.io.InputStream;

import static org.nice.excel.io.FileIteratorConstant.*;

/**
 * <p>
 * 文件信息迭代器静态工厂类
 * </p>
 *
 * @author WECENG
 * @since 2021/7/7 10:30
 */
public class FileEntryIteratorFactory {

    public static FileEntryIterator create(String fileName, InputStream inputStream) throws ExcelException {
        Assert.notNull(fileName, "fileName must not be null");
        Assert.isTrue(fileName.contains(POINT), "fileName should have suffix");
        String fileSuffixName = fileName.substring(fileName.lastIndexOf(POINT) + 1);
        switch (fileSuffixName) {
            case SUFFIX_ZIP:
                return new ZipFileEntryIterator(inputStream);
            case SUFFIX_RAR:
                return new RarFileEntryIterator(inputStream);
            case SUFFIX_XLS:
            case SUFFIX_XLSX:
            case SUFFIX_PDF:
                return new RawFileEntryIterator(new FileEntry(fileName, inputStream));
            default:
                throw new ExcelException("暂不支持该文件格式！{" + fileSuffixName + "}");
        }
    }

}
