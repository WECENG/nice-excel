package org.nice.excel.processor;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.enums.CellExtraTypeEnum;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.listener.ReadListener;
import org.apache.commons.compress.utils.IOUtils;
import org.nice.excel.*;
import org.nice.excel.io.FileEntry;
import org.nice.excel.listener.CustomModelEventListener;
import org.nice.excel.listener.TypeMappingListener;
import org.springframework.util.StringUtils;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * <p>
 * 默认excel处理器
 * </p>
 *
 * @author WECENG
 * @since 2021/7/6 11:02
 */
public class DefaultExcelProcessor implements ExcelProcessor {

    @Override
    public void execute(InputStream in, Class<?> target, List<TypeMappingListener<?>> listenerList) throws ExcelException {
        execute(in, target, listenerList, null);
    }

    @Override
    public void execute(InputStream in, Class<?> target, List<TypeMappingListener<?>> listenerList, Map<Object, Object> customMap) throws ExcelException {
        build(in, target, listenerList).customObject(customMap).doReadAll();
    }

    @Override
    public void execute(InputStream in, Class<?> target, ReadListener<?> listener) throws ExcelException {
        execute(in, target, listener, null);
    }

    @Override
    public void execute(InputStream in, Class<?> target, ReadListener<?> listener, Map<Object, Object> customMap) throws ExcelException {
        build(in, target, listener).customObject(customMap).doReadAll();
    }

    @Override
    public void execute(InputStream in, Class<?> target, ReadListener<?> listener, Map<Object, Object> customMap, String sheetName) throws ExcelException {
        ExcelReaderBuilder builder = build(in, target, listener).customObject(customMap);
        if (StringUtils.isEmpty(sheetName)) {
            builder.doReadAll();
        } else {
            builder.sheet(sheetName).doRead();
        }
    }

    @Override
    public void execute(InputStream in, Class<?> target, ReadListener<?> listener, Map<Object, Object> customMap, String sheetName, Integer rowNum) throws ExcelException {
        ExcelReaderBuilder builder = build(in, target, listener, rowNum).customObject(customMap);
        if (StringUtils.isEmpty(sheetName)) {
            builder.doReadAll();
        } else {
            builder.sheet(sheetName).doRead();
        }
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, List<TypeMappingListener<?>> listenerList) throws ExcelException {
        Class<?> target = modelFunction.applyToModel(fileName);
        execute(in, target, listenerList);
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, List<TypeMappingListener<?>> listenerList, Map<Object, Object> customMap) throws ExcelException {
        Class<?> target = modelFunction.applyToModel(fileName);
        if (Objects.nonNull(target)) {
            execute(in, target, listenerList, customMap);
        }
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction) throws ExcelException {
        Class<?> target = modelFunction.applyToModel(fileName);
        ReadListener<?> readListener = listenerFunction.applyToListener(fileName, target);
        if (Objects.nonNull(target) && Objects.nonNull(readListener)) {
            execute(in, target, readListener);
        }
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction, Map<Object, Object> customMap) throws ExcelException {
        Class<?> target = modelFunction.applyToModel(fileName);
        ReadListener<?> readListener = listenerFunction.applyToListener(fileName, target);
        if (Objects.nonNull(target) && Objects.nonNull(readListener)) {
            execute(in, target, readListener, customMap);
        }
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction, CustomFunction customFunction, Map<Object, Object> customMap) throws ExcelException {
        Class<?> target = modelFunction.applyToModel(fileName);
        ReadListener<?> readListener = listenerFunction.applyToListener(fileName, target);
        customFunction.apply(new FileEntry(fileName, in), customMap);
        if (Objects.nonNull(target) && Objects.nonNull(readListener)) {
            execute(in, target, readListener, customMap);
        }
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction, SheetFunction sheetFunction, CustomFunction customFunction, Map<Object, Object> customMap) throws ExcelException {
        Class<?> target = modelFunction.applyToModel(fileName);
        ReadListener<?> readListener = listenerFunction.applyToListener(fileName, target);
        String sheetName = sheetFunction.applyToSheet(fileName, target);
        customFunction.apply(new FileEntry(fileName, in), customMap);
        if (Objects.nonNull(target) && Objects.nonNull(readListener)) {
            execute(in, target, readListener, customMap, sheetName);
        }
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction, SheetFunction sheetFunction, RowNumFunction rowNumFunction, CustomFunction customFunction, Map<Object, Object> customMap) throws ExcelException {
        try {
            Class<?> target = modelFunction.applyToModel(fileName);
            ReadListener<?> readListener = listenerFunction.applyToListener(fileName, target);
            String sheetName = sheetFunction.applyToSheet(fileName, target);
            Integer rowNum = rowNumFunction.applyToRowNum(fileName, IOUtils.toByteArray(in), sheetName);
            customFunction.apply(new FileEntry(fileName, in), customMap);
            if (Objects.nonNull(target) && Objects.nonNull(readListener)) {
                execute(in, target, readListener, customMap, sheetName, rowNum);
            }
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }

    /**
     * 构建读builder
     *
     * @param in           io stream
     * @param target       目标模型
     * @param listenerList 结果监听器集合
     * @return builder
     */
    protected ExcelReaderBuilder build(InputStream in, Class<?> target, List<TypeMappingListener<?>> listenerList) {
        ExcelReaderBuilder excelReaderBuilder = EasyExcel.read(in, new CustomModelEventListener())
                .head(target)
                .extraRead(CellExtraTypeEnum.MERGE)
                .useDefaultListener(false);
        listenerList.forEach(excelReaderBuilder::registerReadListener);
        return excelReaderBuilder;
    }

    /**
     * 构建读builder
     *
     * @param in     io stream
     * @param target 目标模型
     * @return builder
     */
    protected ExcelReaderBuilder build(InputStream in, Class<?> target, ReadListener<?> listener) {
        return EasyExcel.read(in, new CustomModelEventListener())
                .head(target)
                .extraRead(CellExtraTypeEnum.MERGE)
                .useDefaultListener(false)
                .registerReadListener(listener);
    }

    /**
     * 构建读builder
     *
     * @param in     io stream
     * @param target 目标模型
     * @param rowNum 总行数
     * @return builder
     */
    protected ExcelReaderBuilder build(InputStream in, Class<?> target, ReadListener<?> listener, Integer rowNum) {
        return EasyExcel.read(in, new CustomModelEventListener(rowNum))
                .head(target)
                .extraRead(CellExtraTypeEnum.MERGE)
                .useDefaultListener(false)
                .registerReadListener(listener);
    }

}
