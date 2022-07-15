package org.nice.excel;

import java.io.IOException;

/**
 * <p>
 * 根据文件信息获取excel总行数
 * </p>
 *
 * @author WECENG
 * @since 2021/7/7 17:22
 */
@FunctionalInterface
public interface RowNumFunction {

    /**
     * 获取总行数
     *
     * @param fileName  文件名称
     * @param fileBytes 文件内容数组
     * @param sheetName 名称
     * @return
     */
    Integer applyToRowNum(String fileName, byte[] fileBytes, String sheetName) throws IOException;

}
