package org.nice.excel.processor;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.nice.excel.ExcelException;
import org.nice.excel.listener.*;
import org.nice.excel.model.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.StringUtils;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.nice.excel.io.FileIteratorConstant.SUFFIX_XLS;
import static org.nice.excel.io.FileIteratorConstant.SUFFIX_XLSX;


public class CompressExcelProcessorTest {

    private final Logger log = LoggerFactory.getLogger(this.getClass());

    public static final String SIMPLE_NAME = "simple_name";

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
    public void testExecute() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/all.zip");
        List<TypeMappingListener<?>> listenerList = new ArrayList<>();
        listenerList.add(new GroupByExcelListener());
        listenerList.add(new RowColMergeExcelListener());
        listenerList.add(new RowMergeExcelListener());
        listenerList.add(new BasicExcelListener());
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        excelProcessor.execute("all.zip", in, fileName -> {
            AtomicReference<Class<?>> target = new AtomicReference<>();
            modelMap.forEach((name, clazz) -> {
                if (fileName.substring(fileName.lastIndexOf("/") + 1, fileName.lastIndexOf(".")).equals(name)) {
                    target.set(clazz);
                }
            });
            return target.get();
        }, listenerList);
    }

    @Test
    public void execute() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        assertThrows(IllegalArgumentException.class, () ->
                excelProcessor.execute(in, RowMergeExcel.class, new ArrayList<>(listenerMap.values())));
        excelProcessor.setCompressName("rowMerge.xlsx");
        excelProcessor.execute(in, RowMergeExcel.class, new ArrayList<>(listenerMap.values()));
    }

    @Test
    public void testExecute1() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        Map<Object, Object> customMap = new HashMap<>(1);
        assertThrows(IllegalArgumentException.class, () ->
                excelProcessor.execute(in, RowMergeExcel.class, new ArrayList<>(listenerMap.values()), customMap));
        excelProcessor.setCompressName("rowMerge.xlsx");
        excelProcessor.execute(in, RowMergeExcel.class, new ArrayList<>(listenerMap.values()), customMap);
    }

    @Test
    public void testExecute2() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        assertThrows(IllegalArgumentException.class, () ->
                excelProcessor.execute(in, RowMergeExcel.class, new RowMergeExcelListener()));
        excelProcessor.setCompressName("rowMerge.xlsx");
        excelProcessor.execute(in, RowMergeExcel.class, new RowMergeExcelListener());
    }

    @Test
    public void testExecute3() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        Map<Object, Object> customMap = new HashMap<>(1);
        assertThrows(IllegalArgumentException.class, () ->
                excelProcessor.execute(in, RowMergeExcel.class, new RowMergeExcelListener(), customMap));
        excelProcessor.setCompressName("rowMerge.xlsx");
        excelProcessor.execute(in, RowMergeExcel.class, new RowMergeExcelListener(), customMap);
    }

    @Test
    public void testExecute4() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/all.zip");
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        excelProcessor.execute("all.zip", in, fileName -> {
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
        }, new ArrayList<>(listenerMap.values()));
    }

    @Test
    public void testExecute5() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/all.zip");
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        Map<Object, Object> customMap = new HashMap<>(1);
        excelProcessor.execute("all.zip", in, fileName -> {
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
        }, new ArrayList<>(listenerMap.values()), customMap);
    }

    @Test
    public void testExecute6() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/all.zip");
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        excelProcessor.execute("all.zip", in, fileName -> {
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
        }, (name, clazz) -> listenerMap.get(clazz));
    }

    @Test
    public void testExecute7() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/all.zip");
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        Map<Object, Object> customMap = new HashMap<>(1);
        excelProcessor.execute("all.zip", in, fileName -> {
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
    }

    @Test
    public void testExecute8() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/all.zip");
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        Map<Object, Object> customMap = new HashMap<>(1);
        excelProcessor.execute(
                "all.zip",
                in,
                fileName -> {
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
                },
                (name, clazz) -> listenerMap.get(clazz),
                (fileEntry, cache) -> {
                    String simpleName = fileEntry.getName().substring(fileEntry.getName().lastIndexOf("/") + 1, fileEntry.getName().lastIndexOf("."));
                    cache.put(SIMPLE_NAME, simpleName);
                    log.info("解析文件名称：{}", simpleName);
                },
                customMap);
    }

    @Test
    public void testExecute9() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        Map<Object, Object> customMap = new HashMap<>(1);
        assertThrows(IllegalArgumentException.class, () ->
                excelProcessor.execute(in, RowMergeExcel.class, new RowMergeExcelListener(), customMap, "Sheet1"));
        excelProcessor.setCompressName("rowMerge.xlsx");
        excelProcessor.execute(in, RowMergeExcel.class, new RowMergeExcelListener(), customMap, "Sheet1");
    }

    @Test
    public void testExecuteZip() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/all.zip");
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        Map<Object, Object> customMap = new HashMap<>(1);
        excelProcessor.execute("all.zip", in, fileName -> {
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
    }

    @Test
    public void testExecuteZipWithTotalRow() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/all.zip");
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        Map<Object, Object> customMap = new HashMap<>(1);
        excelProcessor.execute("all.zip", in, fileName -> {
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
                },
                (name, clazz) -> listenerMap.get(clazz),
                (fileName, model) -> null,
                (fileName, fileBytes, sheetName) -> {
                    Workbook workbook;
                    ByteArrayInputStream inputStream = new ByteArrayInputStream(fileBytes);
                    if (fileName.endsWith(SUFFIX_XLS)) {
                        workbook = new HSSFWorkbook(inputStream);
                    } else if (fileName.endsWith(SUFFIX_XLSX)) {
                        workbook = new XSSFWorkbook(inputStream);
                    } else {
                        return null;
                    }
                    Sheet sheet = StringUtils.isEmpty(sheetName) ? workbook.getSheetAt(0) : workbook.getSheet(sheetName);
                    return sheet.getLastRowNum() + 1;
                },
                (fileEntry, model) -> {
                },
                customMap);
    }


    @Test
    public void testExecuteRar() throws ExcelException {
        InputStream in = this.getClass().getClassLoader().getResourceAsStream("dir/all.rar");
        CompressExcelProcessor excelProcessor = new CompressExcelProcessor();
        Map<Object, Object> customMap = new HashMap<>(1);
        excelProcessor.execute("all.rar", in, fileName -> {
            AtomicReference<Class<?>> target = new AtomicReference<>();
            modelMap.forEach((name, clazz) -> {
                if (fileName.substring(fileName.lastIndexOf("\\") + 1, fileName.lastIndexOf(".")).equals(name)) {
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
    }
}