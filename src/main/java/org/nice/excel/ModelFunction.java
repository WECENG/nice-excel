package org.nice.excel;

/**
 * <p>
 * 根据文件名称获取目标模型接口函数
 * </p>
 *
 * @author WECENG
 * @since 2021/7/7 16:26
 */
@FunctionalInterface
public interface ModelFunction {

    /**
     * 获取目标模型
     *
     * @param fileName 文件名称
     * @return 目标模型
     */
    Class<?> applyToModel(String fileName);

}
