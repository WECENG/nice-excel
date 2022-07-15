package org.nice.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import org.nice.excel.utils.ReflectionUtil;

import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

/**
 * <p>
 * 类型映射监听器
 * </p>
 *
 * @author WECENG
 * @since 2021/7/1 15:43
 */
public abstract class TypeMappingListener<T> extends AnalysisEventListener<T> {
    @Override
    public void invoke(T t, AnalysisContext context) {
        Class<?> tClass = ReflectionUtil.getGenericsClass(getClass());
        if (Objects.nonNull(t) && tClass.equals(context.currentReadHolder().excelReadHeadProperty().getHeadClazz())) {
            Map<?, ?> customMap = new HashMap<>(1);
            if (Objects.nonNull(context.getCustom()) && context.getCustom() instanceof Map) {
                customMap = (Map<?, ?>) context.getCustom();
            }
            doInvoke(t, context, customMap);
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        Map<?, ?> customMap = new HashMap<>(1);
        if (Objects.nonNull(context.getCustom()) && context.getCustom() instanceof Map) {
            customMap = (Map<?, ?>) context.getCustom();
        }
        post(context, customMap);
    }

    /**
     * 当excel被解析后，会调用该方法
     *
     * @param t         excel解析后的模型类
     * @param context   excel解析上下文
     * @param customMap 自定义map
     */
    public abstract void doInvoke(T t, AnalysisContext context, Map<?, ?> customMap);

    /**
     * workbook解析完后调用
     *
     * @param context   excel解析上下文
     * @param customMap 自定义map
     */
    public abstract void post(AnalysisContext context, Map<?, ?> customMap);
}
