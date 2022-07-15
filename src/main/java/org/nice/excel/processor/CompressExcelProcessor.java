package org.nice.excel.processor;

import com.alibaba.excel.read.listener.ReadListener;
import org.apache.commons.compress.utils.IOUtils;
import org.nice.excel.*;
import org.nice.excel.io.FileEntry;
import org.nice.excel.io.FileEntryIterator;
import org.nice.excel.io.FileEntryIteratorFactory;
import org.nice.excel.listener.TypeMappingListener;
import org.springframework.util.Assert;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * <p>
 * 压缩包excel处理器
 * </p>
 *
 * @author WECENG
 * @since 2021/7/7 14:54
 */
public class CompressExcelProcessor extends DefaultExcelProcessor {

    private String compressName;

    public String getCompressName() {
        return compressName;
    }

    public void setCompressName(String compressName) {
        this.compressName = compressName;
    }

    @Override
    public void execute(InputStream in, Class<?> target, List<TypeMappingListener<?>> listenerList) throws ExcelException {
        execute(in, target, listenerList, null);
    }

    @Override
    public void execute(InputStream in, Class<?> target, List<TypeMappingListener<?>> listenerList, Map<Object, Object> customMap) throws ExcelException {
        Assert.notNull(compressName, "please set compressName!");
        FileEntryIterator iterator = FileEntryIteratorFactory.create(compressName, in);
        while (iterator.hasNext()) {
            try {
                // avoid affecting IO stream
                byte[] bytes = IOUtils.toByteArray(iterator.next().getExcelInputStream());
                if (0 < bytes.length) {
                    super.execute(new ByteArrayInputStream(bytes), target, listenerList, customMap);
                }
            } catch (IOException e) {
                throw new ExcelException(e);
            }
        }
    }

    @Override
    public void execute(InputStream in, Class<?> target, ReadListener<?> listener) throws ExcelException {
        execute(in, target, listener, null);
    }

    @Override
    public void execute(InputStream in, Class<?> target, ReadListener<?> listener, Map<Object, Object> customMap) throws ExcelException {
        Assert.notNull(compressName, "please set compressName!");
        FileEntryIterator iterator = FileEntryIteratorFactory.create(compressName, in);
        while (iterator.hasNext()) {
            try {
                FileEntry fileEntry = iterator.next();
                // avoid affecting IO stream
                byte[] bytes = IOUtils.toByteArray(fileEntry.getExcelInputStream());
                if (0 < bytes.length) {
                    super.execute(new ByteArrayInputStream(bytes), target, listener, customMap);
                }
            } catch (IOException e) {
                throw new ExcelException(e);
            }
        }
    }

    @Override
    public void execute(InputStream in, Class<?> target, ReadListener<?> listener, Map<Object, Object> customMap, String sheetName) throws ExcelException {
        Assert.notNull(compressName, "please set compressName!");
        FileEntryIterator iterator = FileEntryIteratorFactory.create(compressName, in);
        while (iterator.hasNext()) {
            try {
                FileEntry fileEntry = iterator.next();
                // avoid affecting IO stream
                byte[] bytes = IOUtils.toByteArray(fileEntry.getExcelInputStream());
                if (0 < bytes.length) {
                    super.execute(new ByteArrayInputStream(bytes), target, listener, customMap, sheetName);
                }
            } catch (IOException e) {
                throw new ExcelException(e);
            }
        }
    }

    @Override
    public void execute(InputStream in, Class<?> target, ReadListener<?> listener, Map<Object, Object> customMap, String sheetName, Integer rowNum) throws ExcelException {
        Assert.notNull(compressName, "please set compressName!");
        FileEntryIterator iterator = FileEntryIteratorFactory.create(compressName, in);
        while (iterator.hasNext()) {
            try {
                FileEntry fileEntry = iterator.next();
                // avoid affecting IO stream
                byte[] bytes = IOUtils.toByteArray(fileEntry.getExcelInputStream());
                if (0 < bytes.length) {
                    super.execute(new ByteArrayInputStream(bytes), target, listener, customMap, sheetName, rowNum);
                }
            } catch (IOException e) {
                throw new ExcelException(e);
            }
        }
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, List<TypeMappingListener<?>> listenerList) throws ExcelException {
        execute(fileName, in, modelFunction, listenerList, null);
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, List<TypeMappingListener<?>> listenerList, Map<Object, Object> customMap) throws ExcelException {
        FileEntryIterator iterator = FileEntryIteratorFactory.create(fileName, in);
        while (iterator.hasNext()) {
            try {
                FileEntry fileEntry = iterator.next();
                Class<?> target = modelFunction.applyToModel(fileEntry.getName());
                if (Objects.isNull(target)) {
                    continue;
                }
                // avoid affecting IO stream
                byte[] bytes = IOUtils.toByteArray(fileEntry.getExcelInputStream());
                if (0 < bytes.length) {
                    super.execute(new ByteArrayInputStream(bytes), target, listenerList, customMap);
                }
            } catch (IOException e) {
                throw new ExcelException(e);
            }
        }
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction) throws ExcelException {
        execute(fileName, in, modelFunction, listenerFunction, null);
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction, Map<Object, Object> customMap) throws ExcelException {
        FileEntryIterator iterator = FileEntryIteratorFactory.create(fileName, in);
        while (iterator.hasNext()) {
            try {
                FileEntry fileEntry = iterator.next();
                Class<?> target = modelFunction.applyToModel(fileEntry.getName());
                ReadListener<?> listener = listenerFunction.applyToListener(fileEntry.getName(), target);
                if (Objects.isNull(target) || Objects.isNull(listener)) {
                    continue;
                }
                // avoid affecting IO stream
                byte[] bytes = IOUtils.toByteArray(fileEntry.getExcelInputStream());
                if (0 < bytes.length) {
                    super.execute(new ByteArrayInputStream(bytes), target, listener, customMap);
                }
            } catch (IOException e) {
                throw new ExcelException(e);
            }
        }
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction, CustomFunction customFunction, Map<Object, Object> customMap) throws ExcelException {
        FileEntryIterator iterator = FileEntryIteratorFactory.create(fileName, in);
        while (iterator.hasNext()) {
            try {
                FileEntry fileEntry = iterator.next();
                Class<?> target = modelFunction.applyToModel(fileEntry.getName());
                ReadListener<?> listener = listenerFunction.applyToListener(fileEntry.getName(), target);
                customFunction.apply(fileEntry, customMap);
                if (Objects.isNull(target) || Objects.isNull(listener)) {
                    continue;
                }
                // avoid affecting IO stream
                byte[] bytes = IOUtils.toByteArray(fileEntry.getExcelInputStream());
                if (0 < bytes.length) {
                    super.execute(new ByteArrayInputStream(bytes), target, listener, customMap);
                }
            } catch (IOException e) {
                throw new ExcelException(e);
            }
        }
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction, SheetFunction sheetFunction, CustomFunction customFunction, Map<Object, Object> customMap) throws ExcelException {
        FileEntryIterator iterator = FileEntryIteratorFactory.create(fileName, in);
        while (iterator.hasNext()) {
            try {
                Class<?> target = modelFunction.applyToModel(iterator.next().getName());
                ReadListener<?> listener = listenerFunction.applyToListener(iterator.next().getName(), target);
                String sheetName = sheetFunction.applyToSheet(fileName, target);
                customFunction.apply(iterator.next(), customMap);
                if (Objects.isNull(target) || Objects.isNull(listener)) {
                    continue;
                }
                // avoid affecting IO stream
                byte[] bytes = IOUtils.toByteArray(iterator.next().getExcelInputStream());
                if (0 < bytes.length) {
                    super.execute(new ByteArrayInputStream(bytes), target, listener, customMap, sheetName);
                }
            } catch (IOException e) {
                throw new ExcelException(e);
            }
        }
    }

    @Override
    public void execute(String fileName, InputStream in, ModelFunction modelFunction, ListenerFunction listenerFunction, SheetFunction sheetFunction, RowNumFunction rowNumFunction, CustomFunction customFunction, Map<Object, Object> customMap) throws ExcelException {
        FileEntryIterator iterator = FileEntryIteratorFactory.create(fileName, in);
        while (iterator.hasNext()) {
            try {
                FileEntry fileEntry = iterator.next();
                Class<?> target = modelFunction.applyToModel(fileEntry.getName());
                ReadListener<?> listener = listenerFunction.applyToListener(iterator.next().getName(), target);
                String sheetName = sheetFunction.applyToSheet(fileName, target);
                byte[] bytes = IOUtils.toByteArray(iterator.next().getExcelInputStream());
                Integer rowNum = rowNumFunction.applyToRowNum(fileEntry.getName(), bytes, sheetName);
                customFunction.apply(iterator.next(), customMap);
                if (Objects.isNull(target) || Objects.isNull(listener)) {
                    continue;
                }
                if (0 < bytes.length) {
                    super.execute(new ByteArrayInputStream(bytes), target, listener, customMap, sheetName, rowNum);
                }
            } catch (IOException e) {
                throw new ExcelException(e);
            }
        }
    }
}
