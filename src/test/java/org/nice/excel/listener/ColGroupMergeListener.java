package org.nice.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import org.nice.excel.model.ColGroupMerge;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * <p>
 * 分组后合并监听器
 * </p>
 *
 * @author WECENG
 * @since 2021/7/5 16:26
 */
public class ColGroupMergeListener extends TypeMappingListener<List<ColGroupMerge>> {

    private final Logger log = LoggerFactory.getLogger(this.getClass());

    @Override
    public void doInvoke(List<org.nice.excel.model.ColGroupMerge> colGroupMergeList, AnalysisContext context, Map<?, ?> customMap) {
        log.info("成功解析: {}", colGroupMergeList);
        assertEquals(colGroupMergeList.get(0).getDataList().size(), 96);
    }


    @Override
    public void post(AnalysisContext context, Map<?, ?> customMap) {

    }

    @Override
    public void onException(Exception exception, AnalysisContext context) throws Exception {
        super.onException(exception, context);
    }
}
