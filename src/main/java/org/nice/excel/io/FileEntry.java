package org.nice.excel.io;

import java.io.InputStream;

/**
 * <p>
 * 文件条目
 * </p>
 *
 * @author chenwe@tsintergy.com
 * @date 2021/7/6 11:02
 */
public class FileEntry {

    /**
     * 文件条目名称
     */
    private String name;

    /**
     * 文件条目输入流
     */
    private InputStream excelInputStream;

    public FileEntry(String name, InputStream excelInputStream) {
        this.name = name;
        this.excelInputStream = excelInputStream;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public InputStream getExcelInputStream() {
        return excelInputStream;
    }

    public void setExcelInputStream(InputStream excelInputStream) {
        this.excelInputStream = excelInputStream;
    }
}
