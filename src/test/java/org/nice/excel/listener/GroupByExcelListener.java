package org.nice.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;
import org.assertj.core.util.Lists;
import org.nice.excel.model.GroupByExcel;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertIterableEquals;

/**
 * <p>
 * 分组excel监听器
 * </p>
 *
 * @author WECENG
 * @since 2021/7/5 14:45
 */
public class GroupByExcelListener extends TypeMappingListener<List<GroupByExcel>> {

    private final Logger log = LoggerFactory.getLogger(this.getClass());

    @Override
    public void doInvoke(List<GroupByExcel> groupByExcels, AnalysisContext context, Map<?, ?> customMap) {
        log.info("成功解析: {}", groupByExcels);
        assertEquals(8, groupByExcels.size());
        assertIterableEquals(Lists.list(2, 2, 2, 2, 2, 2, 2, 2),
                groupByExcels.stream().map(GroupByExcel::getDataList).map(List::size).collect(Collectors.toList()));
    }

    @Override
    public void invokeHead(Map<Integer, CellData> headMap, AnalysisContext context) {

    }

    @Override
    public void post(AnalysisContext context, Map<?, ?> customMap) {

    }
}
