package org.nice.excel.exception;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.exception.ExcelDataConvertException;
import com.alibaba.excel.metadata.CellData;
import org.junit.jupiter.api.Test;
import org.nice.excel.listener.CustomModelEventListener;
import org.nice.excel.listener.TypeMappingListener;
import org.nice.excel.model.RowMergeExcel;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * <p>
 * 单元格数据解析错误异常单元测试类
 * </p>
 *
 * @author WECENG
 * @since 2021/7/20 9:46
 */
public class DataExceptionTest {

    @Test
    public void test() {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("exception/exception.xlsx");
        EasyExcel.read(in, new CustomModelEventListener())
                .useDefaultListener(false)
                .registerReadListener(new CustomListener())
                .head(RowMergeExcel.class)
                .doReadAll();
    }

    public static class CustomListener extends TypeMappingListener<RowMergeExcel> {

        private final Logger log = LoggerFactory.getLogger(this.getClass());

        @Override
        public void doInvoke(RowMergeExcel rowMergeExcel, AnalysisContext context, Map<?, ?> customMap) {

        }

        @Override
        public void invokeHead(Map<Integer, CellData> headMap, AnalysisContext context) {

        }

        @Override
        public void post(AnalysisContext context, Map<?, ?> customMap) {

        }

        @Override
        public void onException(Exception exception, AnalysisContext context) throws Exception {
            assertEquals(ExcelDataConvertException.class, exception.getClass());
            if (exception instanceof ExcelDataConvertException) {
                ExcelDataConvertException e = (ExcelDataConvertException) exception;
                log.info("解析异常，错误行数：{}，错误列数：{}，错误内容：{}", e.getRowIndex() + 1, e.getColumnIndex() + 1, e.getCellData().toString());
            }
        }
    }

}
