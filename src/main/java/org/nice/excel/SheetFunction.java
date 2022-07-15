package org.nice.excel;

/**
 * <p>
 * 根据文件名称和目标模型获取sheet名称接口函数
 * </p>
 *
 * @author WECENG
 * @since 2021/7/7 17:22
 */
@FunctionalInterface
public interface SheetFunction {

    /**
     * 获取sheet名称
     *
     * @param fileName 文件名称
     * @param model    目标模型
     * @return
     */
    String applyToSheet(String fileName, Class<?> model);

}
