package org.nice.excel;


import org.nice.excel.io.FileEntry;

import java.util.Map;

/**
 * <p>
 * 用于处理自定义逻辑函数接口
 * </p>
 *
 * @author WECENG
 * @since 2021/7/20 16:44
 */
@FunctionalInterface
public interface CustomFunction {

    /**
     * 自定义逻辑
     *
     * @param fileEntry 文件信息
     * @param customMap 自定义缓存
     */
    void apply(FileEntry fileEntry, Map<Object, Object> customMap);

}
