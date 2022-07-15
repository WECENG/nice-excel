package org.nice.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;
import org.nice.excel.model.ColMergeExcel;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * <p>
 * 按列合并excel监听器
 * </p>
 *
 * @author WECENG
 * @since 2022/7/15 10:25
 */
public class ColMergeExcelListener extends TypeMappingListener<ColMergeExcel> {

    private final Logger log = LoggerFactory.getLogger(this.getClass());

    @Override
    public void doInvoke(ColMergeExcel colMergeExcel, AnalysisContext context, Map<?, ?> customMap) {
        log.info("成功解析: {}", colMergeExcel);
        assertEquals(96, colMergeExcel.getDataList().size());
    }

    @Override
    public void invokeHead(Map<Integer, CellData> headMap, AnalysisContext context) {
    }

    @Override
    public void post(AnalysisContext context, Map<?, ?> customMap) {

    }

}
