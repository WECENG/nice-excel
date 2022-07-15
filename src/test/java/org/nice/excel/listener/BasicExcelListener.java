package org.nice.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;
import org.nice.excel.model.BasicExcel;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertNotNull;

/**
 * <p>
 * 基础excel监听器
 * </p>
 *
 * @author WECENG
 * @since 2021/7/5 16:11
 */
public class BasicExcelListener extends TypeMappingListener<BasicExcel> {

    private final Logger log = LoggerFactory.getLogger(this.getClass());

    @Override
    public void doInvoke(BasicExcel basicExcel, AnalysisContext context, Map<?, ?> customMap) {
        log.info("成功解析: {}", basicExcel);
        assertNotNull(basicExcel.getNumber());
        assertNotNull(basicExcel.getStr());
    }

    @Override
    public void invokeHead(Map<Integer, CellData> headMap, AnalysisContext context) {

    }

    @Override
    public void post(AnalysisContext context, Map<?, ?> customMap) {

    }
}
