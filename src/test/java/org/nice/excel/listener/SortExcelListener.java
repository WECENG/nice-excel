package org.nice.excel.listener;

import cn.hutool.core.date.DateUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;
import org.nice.excel.model.RowMergeExcel;
import org.nice.excel.model.SortExcel;
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
public class SortExcelListener extends TypeMappingListener<SortExcel> {

    private final Logger log = LoggerFactory.getLogger(this.getClass());

    @Override
    public void doInvoke(SortExcel sortExcel, AnalysisContext context, Map<?, ?> customMap) {
        log.info("成功解析: {}", sortExcel);
        //assert sort
        assertEquals(sortExcel.getDataList().get(0), BigDecimal.valueOf(1536.36));
    }

    @Override
    public void invokeHead(Map<Integer, CellData> headMap, AnalysisContext context) {

    }

    @Override
    public void post(AnalysisContext context, Map<?, ?> customMap) {

    }
}
