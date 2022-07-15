package org.nice.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;
import org.nice.excel.model.RowColMergeExcel;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

/**
 * <p>
 * 行列合并excel监听器
 * </p>
 *
 * @author WECENG
 * @since 2021/7/5 16:35
 */
public class RowColMergeExcelListener extends TypeMappingListener<RowColMergeExcel> {

    private final Logger log = LoggerFactory.getLogger(this.getClass());

    @Override
    public void doInvoke(RowColMergeExcel rowColMergeExcel, AnalysisContext context, Map<?, ?> customMap) {
        log.info("成功解析: {}", rowColMergeExcel);
        assertNotNull(rowColMergeExcel);
        assertEquals(rowColMergeExcel.getDataList().size(), 96);
    }

    @Override
    public void invokeHead(Map<Integer, CellData> headMap, AnalysisContext context) {

    }

    @Override
    public void post(AnalysisContext context, Map<?, ?> customMap) {

    }
}
