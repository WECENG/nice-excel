package org.nice.excel;

import com.alibaba.excel.read.listener.ReadListener;

/**
 * <p>
 * 根据文件名称和目标模型获取结果监听器接口函数
 * </p>
 *
 * @author WECENG
 * @since 2021/7/7 17:22
 */
@FunctionalInterface
public interface ListenerFunction {

    /**
     * 获取结果监听器
     *
     * @param fileName 文件名称
     * @param model    目标模型
     * @return
     */
    ReadListener<?> applyToListener(String fileName, Class<?> model);

}
