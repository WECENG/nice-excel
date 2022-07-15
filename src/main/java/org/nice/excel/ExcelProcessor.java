package org.nice.excel;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.listener.ReadListener;
import org.nice.excel.listener.TypeMappingListener;

import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * <p>
 * excel处理器
 * </p>
 *
 * @author WECENG
 * @since 2021/7/6 11:02
 */
public interface ExcelProcessor {

    /**
     * excel解析
     *
     * @param in           io stream
     * @param target       目标模型
     * @param listenerList 结果监听器集合
     * @see TypeMappingListener
     */
    void execute(InputStream in, Class<?> target, List<TypeMappingListener<?>> listenerList) throws ExcelException;

    /**
     * excel解析
     *
     * @param in           io stream
     * @param target       目标模型
     * @param listenerList 监听器集合
     * @param customMap    This object can be read in the Listener {@link AnalysisEventListener#invoke(Object, AnalysisContext)}
     *                     {@link AnalysisContext#getCustom()}
     * @see TypeMappingListener
     */
    void execute(InputStream in, Class<?> target, List<TypeMappingListener<?>> listenerList, Map<Object, Object> customMap) throws ExcelException;

    /**
     * excel解析
     *
     * @param in       io stream
     * @param target   目标模型
     * @param listener 结果监听器
     */
    void execute(InputStream in, Class<?> target, ReadListener<?> listener) throws ExcelException;

    /**
     * excel解析
     *
     * @param in        io stream
     * @param target    目标模型
     * @param listener  结果监听器
     * @param customMap This object can be read in the Listener {@link AnalysisEventListener#invoke(Object, AnalysisContext)}
     *                  {@link AnalysisContext#getCustom()}
     */
    void execute(InputStream in, Class<?> target, ReadListener<?> listener, Map<Object, Object> customMap) throws ExcelException;

    /**
     * excel解析
     *
     * @param in        io stream
     * @param target    目标模型
     * @param listener  结果监听器
     * @param customMap This object can be read in the Listener {@link AnalysisEventListener#invoke(Object, AnalysisContext)}
     *                  {@link AnalysisContext#getCustom()}
     * @param sheetName sheet名称
     */
    void execute(InputStream in, Class<?> target, ReadListener<?> listener, Map<Object, Object> customMap, String sheetName) throws ExcelException;

    /**
     * excel解析
     *
     * @param in        io stream
     * @param target    目标模型
     * @param listener  结果监听器
     * @param customMap This object can be read in the Listener {@link AnalysisEventListener#invoke(Object, AnalysisContext)}
     *                  {@link AnalysisContext#getCustom()}
     * @param sheetName sheet名称
     * @param rowNum    总行数
     */
    void execute(InputStream in, Class<?> target, ReadListener<?> listener, Map<Object, Object> customMap, String sheetName, Integer rowNum) throws ExcelException;

    /**
     * excel解析
     *
     * @param fileName      文件名称
     * @param in            io stream
     * @param modelFunction 获取目标模型
     * @param listenerList  结果监听器集合
     * @see TypeMappingListener
     */
    void execute(String fileName, InputStream in, ModelFunction modelFunction, List<TypeMappingListener<?>> listenerList) throws ExcelException;

    /**
     * excel解析
     *
     * @param fileName      文件名称
     * @param in            io stream
     * @param modelFunction 获取目标模型
     * @param listenerList  监听器集合
     * @param customMap     This object can be read in the Listener {@link AnalysisEventListener#invoke(Object, AnalysisContext)}
     *                      {@link AnalysisContext#getCustom()}
     * @see TypeMappingListener
     */
    void execute(String fileName, InputStream in, ModelFunction modelFunction, List<TypeMappingListener<?>> listenerList, Map<Object, Object> customMap) throws ExcelException;

    /**
     * excel解析
     *
     * @param fileName         文件名称
     * @param in               io stream
     * @param modelFunction    目标模型
     * @param listenerFunction 结果监听器
     */
    void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction) throws ExcelException;

    /**
     * excel解析
     *
     * @param fileName         文件名称
     * @param in               io stream
     * @param modelFunction    目标模型
     * @param listenerFunction 结果监听器
     * @param customMap        This object can be read in the Listener {@link AnalysisEventListener#invoke(Object, AnalysisContext)}
     *                         {@link AnalysisContext#getCustom()}
     */
    void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction, Map<Object, Object> customMap) throws ExcelException;

    /**
     * excel解析
     *
     * @param fileName         文件名称
     * @param in               io stream
     * @param modelFunction    目标模型
     * @param listenerFunction 结果监听器
     * @param customFunction   自定义解析
     * @param customMap        This object can be read in the Listener {@link AnalysisEventListener#invoke(Object, AnalysisContext)}
     *                         {@link AnalysisContext#getCustom()}
     */
    void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction, CustomFunction customFunction, Map<Object, Object> customMap) throws ExcelException;

    /**
     * excel解析
     *
     * @param fileName         文件名称
     * @param in               io stream
     * @param modelFunction    目标模型
     * @param listenerFunction 结果监听器
     * @param customFunction   自定义解析
     * @param sheetFunction    sheet名称解析
     * @param customMap        This object can be read in the Listener {@link AnalysisEventListener#invoke(Object, AnalysisContext)}
     *                         {@link AnalysisContext#getCustom()}
     */
    void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction, SheetFunction sheetFunction, CustomFunction customFunction, Map<Object, Object> customMap) throws ExcelException;


    /**
     * excel解析
     *
     * @param fileName         文件名称
     * @param in               io stream
     * @param modelFunction    目标模型
     * @param listenerFunction 结果监听器
     * @param sheetFunction    sheet名称解析
     * @param rowNumFunction   总行数解析
     * @param customFunction   自定义解析
     * @param customMap        This object can be read in the Listener {@link AnalysisEventListener#invoke(Object, AnalysisContext)}
     *                         {@link AnalysisContext#getCustom()}
     */
    void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction, SheetFunction sheetFunction, RowNumFunction rowNumFunction, CustomFunction customFunction, Map<Object, Object> customMap) throws ExcelException;

}
