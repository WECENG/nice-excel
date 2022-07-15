package org.nice.excel.listener;

import cn.hutool.core.date.DateUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;
import org.nice.excel.model.RowMergeExcel;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigDecimal;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * <p>
 * 按行合并excel监听器
 * </p>
 *
 * @author WECENG
 * @since 2021/7/5 10:35
 */
public class RowMergeExcelListener extends TypeMappingListener<RowMergeExcel> {

    private final Logger log = LoggerFactory.getLogger(this.getClass());

    @Override
    public void doInvoke(RowMergeExcel rowMergeExcel, AnalysisContext context, Map<?, ?> customMap) {
        log.info("成功解析: {}", rowMergeExcel);
        assertTrue(DateUtil.isSameDay(rowMergeExcel.getDate(), DateUtil.parse("2021-04-26", "yyyy-MM-dd")));
        assertEquals(96, rowMergeExcel.getDataList().size());
        //assert sort
        assertEquals(rowMergeExcel.getDataList().get(0), BigDecimal.valueOf(1536.36));
    }

    @Override
    public void invokeHead(Map<Integer, CellData> headMap, AnalysisContext context) {

    }

    @Override
    public void post(AnalysisContext context, Map<?, ?> customMap) {

    }
}
