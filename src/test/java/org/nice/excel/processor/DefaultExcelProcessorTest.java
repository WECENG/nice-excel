package org.nice.excel.processor;

import cn.hutool.core.date.DateUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;
import org.junit.jupiter.api.Test;
import org.nice.excel.ExcelException;
import org.nice.excel.listener.*;
import org.nice.excel.model.RowMergeExcel;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

public class DefaultExcelProcessorTest {

    private final Logger log = LoggerFactory.getLogger(this.getClass());

    public static final String KEY = "defaultExcelProcessor";

    public static final String SIMPLE_NAME = "simple_name";

    @Test
    public void testExecute() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        List<TypeMappingListener<?>> listenerList = new ArrayList<>();
        listenerList.add(new GroupByExcelListener());
        listenerList.add(new RowMergeListener());
        listenerList.add(new RowColMergeExcelListener());
        DefaultExcelProcessor processor = new DefaultExcelProcessor();
        processor.execute(in, RowMergeExcel.class, listenerList);
    }

    @Test
    public void testExecute1() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        List<TypeMappingListener<?>> listenerList = new ArrayList<>();
        listenerList.add(new GroupByExcelListener());
        listenerList.add(new RowMergeListener());
        listenerList.add(new RowColMergeExcelListener());
        DefaultExcelProcessor processor = new DefaultExcelProcessor();
        Map<Object, Object> customMap = new HashMap<>(8);
        customMap.put(KEY, processor);
        processor.execute(in, RowMergeExcel.class, listenerList, customMap);
    }

    @Test
    public void testExecute2() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        DefaultExcelProcessor processor = new DefaultExcelProcessor();
        processor.execute(in, RowMergeExcel.class, new RowMergeListener());
    }

    @Test
    public void testExecute3() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        DefaultExcelProcessor processor = new DefaultExcelProcessor();
        Map<Object, Object> customMap = new HashMap<>(8);
        customMap.put(KEY, processor);
        processor.execute(in, RowMergeExcel.class, new RowMergeListener(), customMap);
    }

    @Test
    public void testExecute4() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        List<TypeMappingListener<?>> listenerList = new ArrayList<>();
        listenerList.add(new GroupByExcelListener());
        listenerList.add(new RowMergeListener());
        listenerList.add(new RowColMergeExcelListener());
        DefaultExcelProcessor processor = new DefaultExcelProcessor();
        Map<String, Class<?>> modelMap = new HashMap<>();
        modelMap.put("rowMerge.xlsx", RowMergeExcel.class);
        processor.execute("rowMerge.xlsx", in, modelMap::get, listenerList);
    }

    @Test
    public void testExecute5() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        List<TypeMappingListener<?>> listenerList = new ArrayList<>();
        listenerList.add(new GroupByExcelListener());
        listenerList.add(new RowMergeListener());
        listenerList.add(new RowColMergeExcelListener());
        DefaultExcelProcessor processor = new DefaultExcelProcessor();
        Map<String, Class<?>> modelMap = new HashMap<>();
        modelMap.put("rowMerge.xlsx", RowMergeExcel.class);
        Map<Object, Object> customMap = new HashMap<>(8);
        customMap.put(KEY, processor);
        processor.execute("rowMerge.xlsx", in, modelMap::get, listenerList, customMap);
    }

    @Test
    public void testExecute6() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        DefaultExcelProcessor processor = new DefaultExcelProcessor();
        Map<String, Class<?>> modelMap = new HashMap<>();
        Map<Class<?>, TypeMappingListener<?>> listenerMap = new HashMap<>();
        listenerMap.put(RowMergeExcel.class, new RowMergeExcelListener());
        modelMap.put("rowMerge.xlsx", RowMergeExcel.class);
        processor.execute("rowMerge.xlsx", in, modelMap::get, (fileName, model) -> listenerMap.get(model));
    }

    @Test
    public void testExecute7() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        DefaultExcelProcessor processor = new DefaultExcelProcessor();
        Map<String, Class<?>> modelMap = new HashMap<>();
        Map<Class<?>, TypeMappingListener<?>> listenerMap = new HashMap<>();
        listenerMap.put(RowMergeExcel.class, new RowMergeExcelListener());
        modelMap.put("rowMerge.xlsx", RowMergeExcel.class);
        Map<Object, Object> customMap = new HashMap<>(8);
        customMap.put(KEY, processor);
        processor.execute("rowMerge.xlsx", in, modelMap::get, (fileName, model) -> listenerMap.get(model), customMap);
    }

    @Test
    public void testExecute8() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        DefaultExcelProcessor processor = new DefaultExcelProcessor();
        Map<String, Class<?>> modelMap = new HashMap<>();
        Map<Class<?>, TypeMappingListener<?>> listenerMap = new HashMap<>();
        listenerMap.put(RowMergeExcel.class, new TypeMappingListener<RowMergeExcel>() {
            @Override
            public void doInvoke(RowMergeExcel rowMergeExcel, AnalysisContext context, Map<?, ?> customMap) {
                log.info("成功解析: {}", rowMergeExcel);
                assertTrue(DateUtil.isSameDay(rowMergeExcel.getDate(), DateUtil.parse("2021-04-26", "yyyy-MM-dd")));
                assertEquals(96, rowMergeExcel.getDataList().size());
                assertNotNull(customMap.get(SIMPLE_NAME));
                assertNotNull(customMap.get(KEY));
            }

            @Override
            public void post(AnalysisContext context, Map<?, ?> customMap) {

            }
        });
        modelMap.put("rowMerge.xlsx", RowMergeExcel.class);
        Map<Object, Object> customMap = new HashMap<>(8);
        customMap.put(KEY, processor);
        processor.execute("rowMerge.xlsx", in, modelMap::get, (fileName, model) -> listenerMap.get(model), (fileEntry, cache) -> {
            String simpleName = fileEntry.getName().substring(fileEntry.getName().lastIndexOf("/") + 1, fileEntry.getName().lastIndexOf("."));
            cache.put(SIMPLE_NAME, simpleName);
            log.info("解析文件名称：{}", simpleName);
        }, customMap);
    }

    public static class RowMergeListener extends TypeMappingListener<RowMergeExcel> {

        private final Logger log = LoggerFactory.getLogger(this.getClass());

        @Override
        public void doInvoke(RowMergeExcel rowMergeExcel, AnalysisContext context, Map<?, ?> customMap) {
            log.info("成功解析: {}", rowMergeExcel);
            assertTrue(DateUtil.isSameDay(rowMergeExcel.getDate(), DateUtil.parse("2021-04-26", "yyyy-MM-dd")));
            assertEquals(96, rowMergeExcel.getDataList().size());
        }

        @Override
        public void invokeHead(Map<Integer, CellData> headMap, AnalysisContext context) {

        }

        @Override
        public void post(AnalysisContext context, Map<?, ?> customMap) {

        }
    }


}