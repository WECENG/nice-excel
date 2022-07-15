package org.nice.excel.read;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;
import org.apache.commons.compress.utils.IOUtils;
import org.junit.jupiter.api.Test;
import org.nice.excel.ExcelException;
import org.nice.excel.io.FileEntry;
import org.nice.excel.io.FileEntryIterator;
import org.nice.excel.io.FileEntryIteratorFactory;
import org.nice.excel.listener.*;
import org.nice.excel.model.*;
import org.nice.excel.processor.DefaultExcelProcessor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * <p>
 * 文件名称解析单元测试类
 * </p>
 *
 * @author WECENG
 * @since 2021/7/20 10:27
 */
public class FileNameReadTest {

    private final Logger log = LoggerFactory.getLogger(this.getClass());

    public static final String SIMPLE_NAME = "simple_name";

    public static final String DATE = "date";

    private final static Map<String, Class<?>> modelMap = new HashMap<>();

    private final static Map<Class<?>, TypeMappingListener<?>> listenerMap = new HashMap<>();

    static {
        modelMap.put("basicExcel", BasicExcel.class);
        modelMap.put("rowMerge", RowMergeExcel.class);
        modelMap.put("colMerge", ColMergeExcel.class);
        modelMap.put("rowColMerge", RowColMergeExcel.class);
        modelMap.put("groupBy", GroupByExcel.class);
        modelMap.put("sort", SortExcel.class);

        listenerMap.put(BasicExcel.class, new BasicExcelListener());
        listenerMap.put(RowMergeExcel.class, new RowMergeExcelListener());
        listenerMap.put(ColMergeExcel.class, new ColMergeExcelListener());
        listenerMap.put(RowColMergeExcel.class, new RowColMergeExcelListener());
        listenerMap.put(GroupByExcel.class, new GroupByExcelListener());
        listenerMap.put(SortExcel.class, new SortExcelListener());
    }

    @Test
    public void test() throws ExcelException {
        String fileName = "colMerge.xls";
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/" + fileName);
        DefaultExcelProcessor processor = new DefaultExcelProcessor();
        Map<Object, Object> customMap = new HashMap<>();
        customMap.put(SIMPLE_NAME, "colMerge");
        processor.execute(in, ColMergeExcel.class, new DataSupplyDemandListener(), customMap);
    }

    public static class DataSupplyDemandListener extends TypeMappingListener<ColMergeExcel> {

        private final Logger log = LoggerFactory.getLogger(this.getClass());

        @Override
        public void doInvoke(ColMergeExcel colMergeExcel, AnalysisContext context, Map<?, ?> customMap) {
            log.info("成功解析: {}", colMergeExcel);
            assertEquals(96, colMergeExcel.getDataList().size());
            assertEquals("colMerge", customMap.get(SIMPLE_NAME));
        }

        @Override
        public void invokeHead(Map<Integer, CellData> headMap, AnalysisContext context) {
        }

        @Override
        public void post(AnalysisContext context, Map<?, ?> customMap) {

        }
    }

    @Test
    public void testCompress() throws IOException, ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/all.zip");
        FileEntryIterator fileEntryIterator = FileEntryIteratorFactory.create("all.zip", in);
        while (fileEntryIterator.hasNext()) {
            FileEntry fileEntry = fileEntryIterator.next();
            DefaultExcelProcessor processor = new DefaultExcelProcessor();
            Map<Object, Object> customMap = new HashMap<>();
            customMap.put(SIMPLE_NAME, fileEntry.getName().substring(fileEntry.getName().lastIndexOf("/") + 1, fileEntry.getName().lastIndexOf(".")));
            // avoid affecting IO stream
            byte[] bytes = IOUtils.toByteArray(fileEntry.getExcelInputStream());
            processor.execute(fileEntry.getName(), new ByteArrayInputStream(bytes), fileName -> {
                AtomicReference<Class<?>> target = new AtomicReference<>();
                modelMap.forEach((name, clazz) -> {
                    if (fileName.substring(fileName.lastIndexOf("/") + 1, fileName.lastIndexOf(".")).equals(name)) {
                        target.set(clazz);
                    }
                });
                if (Objects.isNull(target.get())) {
                    modelMap.forEach((name, clazz) -> {
                        if (fileName.contains(name)) {
                            target.set(clazz);
                        }
                    });
                }
                return target.get();
            }, (name, clazz) -> listenerMap.get(clazz), customMap);
            log.info("解析文件名称：{}", customMap.get(SIMPLE_NAME));
        }
    }

}
