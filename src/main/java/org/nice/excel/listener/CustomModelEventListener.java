package org.nice.excel.listener;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.enums.HeadKindEnum;
import com.alibaba.excel.exception.ExcelDataConvertException;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.alibaba.excel.read.listener.ModelBuildEventListener;
import com.alibaba.excel.read.metadata.holder.ReadHolder;
import com.alibaba.excel.read.metadata.property.ExcelReadHeadProperty;
import com.alibaba.excel.util.ConverterUtils;
import net.sf.cglib.beans.BeanMap;
import org.nice.excel.ExcelSortModel;
import org.nice.excel.annotation.ExcelGroup;
import org.nice.excel.annotation.ExcelMerge;
import org.nice.excel.annotation.ExcelSort;
import org.nice.excel.utils.ConvertUtil;
import org.springframework.util.CollectionUtils;

import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * <p>
 * 自定义模型转换器,支持数据分组、合并。
 *
 * @author WECENG
 * @see ExcelGroup
 * @see ExcelMerge
 * @see ExcelSort
 * </p>
 * @since 2021/6/30 17:26
 */
@SuppressWarnings("all")
public class CustomModelEventListener extends ModelBuildEventListener {

    private final Map<Integer, Map<Integer, CellData>> cache = new HashMap<>(8);

    private Integer totalRow;

    public CustomModelEventListener() {
    }

    public CustomModelEventListener(Integer totalRow) {
        this.totalRow = totalRow;
    }

    @Override
    public void invoke(Map<Integer, CellData> cellDataMap, AnalysisContext context) {
        ReadHolder currentReadHolder = context.currentReadHolder();
        if (HeadKindEnum.CLASS.equals(currentReadHolder.excelReadHeadProperty().getHeadKind())) {
            context.readRowHolder()
                    .setCurrentRowAnalysisResult(buildUserModel(cellDataMap, currentReadHolder, context));
            return;
        }
        context.readRowHolder().setCurrentRowAnalysisResult(buildStringList(cellDataMap, currentReadHolder, context));
    }

    private Object buildUserModel(Map<Integer, CellData> cellDataMap, ReadHolder currentReadHolder,
                                  AnalysisContext context) {
        ExcelReadHeadProperty excelReadHeadProperty = currentReadHolder.excelReadHeadProperty();
        Object resultModel;
        try {
            resultModel = excelReadHeadProperty.getHeadClazz().newInstance();
        } catch (Exception e) {
            throw new ExcelDataConvertException(context.readRowHolder().getRowIndex(), 0,
                    new CellData(CellDataTypeEnum.EMPTY), null,
                    "Can not instance class: " + excelReadHeadProperty.getHeadClazz().getName(), e);
        }
        Map<String, Object> fieldMap = new HashMap<>(8);
        Map<Integer, ExcelContentProperty> contentPropertyMap = excelReadHeadProperty.getContentPropertyMap();
        for (Map.Entry<Integer, ExcelContentProperty> entry : contentPropertyMap.entrySet()) {
            Integer index = entry.getKey();
            totalRow = dealCellContent(cellDataMap, context, currentReadHolder, contentPropertyMap, index, fieldMap, totalRow);
        }
        if (!CollectionUtils.isEmpty(cache) && context.readRowHolder().getRowIndex() + 1 != totalRow) {
            return null;
        }
        reset();
        List<Object> dataGroupBy = dealDataGroupBy(excelReadHeadProperty, fieldMap, context);
        if (!CollectionUtils.isEmpty(dataGroupBy)) {
            return dataGroupBy;
        }
        //deal sort
        fieldMap.forEach((k, v) -> {
            if (v instanceof List) {
                if (((List) v).stream().allMatch(itemObj -> Objects.nonNull(itemObj) && itemObj.getClass().equals(ExcelSortModel.class))) {
                    v = ((List<ExcelSortModel>) v).stream().sorted(Comparator.comparing(ExcelSortModel::getSortKey)).map(ExcelSortModel::getSortValue).collect(Collectors.toList());
                    fieldMap.put(k, v);
                }
            }
        });
        BeanMap.create(resultModel).putAll(fieldMap);
        return resultModel;
    }

    /**
     * 重置缓存
     */
    private void reset() {
        cache.clear();
        totalRow = null;
    }

    /**
     * 处理数据分组
     *
     * @param excelReadHeadProperty
     * @param fieldMap
     * @param context
     * @return 分组结果
     */
    private List<Object> dealDataGroupBy(ExcelReadHeadProperty excelReadHeadProperty, Map<String, Object> fieldMap, AnalysisContext context) {
        ExcelGroup excelGroup = (ExcelGroup) Arrays.stream(excelReadHeadProperty.getHeadClazz().getAnnotations()).filter(item -> item.annotationType().equals(ExcelGroup.class)).findAny().orElse(null);
        List<Object> resultList = new ArrayList<>();
        if (Objects.nonNull(excelGroup)) {
            List<Map<String, Object>> groupMapList = new ArrayList<>();
            Map<String, List<List<Integer>>> fieldValGroupMap = new LinkedHashMap<>();
            for (String field : excelGroup.fields()) {
                Object fieldVal = fieldMap.get(field);
                if (Objects.nonNull(fieldVal) && fieldVal instanceof List) {
                    List<Object> fieldValList = (List) fieldVal;
                    Map<Integer, Object> valMap = new HashMap(8);
                    for (int i = 0; i < fieldValList.size(); i++) {
                        valMap.put(i, fieldValList.get(i));
                    }
                    Set<Object> groupKey = fieldValList.stream().filter(Objects::nonNull).collect(Collectors.groupingBy(Function.identity())).keySet();
                    List<List<Integer>> idxList = new ArrayList<>();
                    groupKey.forEach(val -> {
                        List<Integer> itemIdxList = ConvertUtil.collectDataByUdx(valMap, val);
                        idxList.add(itemIdxList.stream().sorted().collect(Collectors.toList()));
                    });
                    fieldValGroupMap.put(field, idxList);
                }
            }
            List<List<Integer>> groupIdxList = new ArrayList<>();
            if (!fieldValGroupMap.isEmpty()) {
                if (fieldValGroupMap.size() == 1) {
                    groupIdxList = fieldValGroupMap.get(excelGroup.fields()[0]);
                } else {
                    groupIdxList = fieldValGroupMap.get(excelGroup.fields()[0]);
                    for (int i = 1; i < fieldValGroupMap.size(); i++) {
                        List<List<Integer>> itemIdxList = new ArrayList<>();
                        for (List<Integer> firstIdx : groupIdxList) {
                            for (List<Integer> secondIdx : fieldValGroupMap.get(fieldValGroupMap.keySet().toArray()[i])) {
                                List<Integer> intersection = (List<Integer>) CollUtil.intersection(firstIdx, secondIdx);
                                if (!CollectionUtils.isEmpty(intersection)) {
                                    itemIdxList.add(intersection.stream().sorted().collect(Collectors.toList()));
                                }
                            }
                        }
                        groupIdxList = itemIdxList.stream().distinct().collect(Collectors.toList());
                    }
                }
                //group data
                groupIdxList.forEach(itemIdxList -> {
                    Map<String, Object> itemMap = new HashMap<>(8);
                    fieldMap.forEach((k, v) -> {
                        itemMap.put(k, ConvertUtil.matchData(itemIdxList, (List) v));
                    });
                    groupMapList.add(itemMap);
                });
                //remove null item
                groupMapList.forEach(item -> {
                    item.forEach((k, v) -> {
                        if (v instanceof List) {
                            ((List) v).removeIf(obj -> CellDataTypeEnum.EMPTY == obj || null == obj);
                        }
                    });
                });
                //deal sort
                groupMapList.forEach(item -> {
                    item.forEach((k, v) -> {
                        if (v instanceof List) {
                            if (((List) v).stream().allMatch(itemObj -> Objects.nonNull(itemObj) && itemObj.getClass().equals(ExcelSortModel.class))) {
                                v = ((List<ExcelSortModel>) v).stream().sorted(Comparator.comparing(ExcelSortModel::getSortKey)).map(ExcelSortModel::getSortValue).collect(Collectors.toList());
                                item.put(k, v);
                            }
                        }
                    });
                });
                //deal group field list to one
                fieldValGroupMap.forEach((k, v) -> groupMapList.forEach(item -> {
                    Object fieldVal = item.get(k);
                    if (fieldVal instanceof List) {
                        item.put(k, ((List) fieldVal).stream().filter(Objects::nonNull).findAny().orElse(null));
                    }
                }));
            }
            groupMapList.forEach(item -> {
                Object model;
                try {
                    model = excelReadHeadProperty.getHeadClazz().newInstance();
                } catch (Exception e) {
                    throw new ExcelDataConvertException(context.readRowHolder().getRowIndex(), 0,
                            new CellData(CellDataTypeEnum.EMPTY), null,
                            "Can not instance class: " + excelReadHeadProperty.getHeadClazz().getName(), e);
                }
                BeanMap.create(model).putAll(item);
                resultList.add(model);
            });
        }
        return resultList;
    }

    /**
     * 处理单元格内容
     *
     * @param cellDataMap
     * @param context
     * @param currentReadHolder
     * @param contentPropertyMap
     * @param propertyIdx
     * @param fieldMap
     * @param totalRow           总行数
     * @return 总行数
     */
    private Integer dealCellContent(Map<Integer, CellData> cellDataMap, AnalysisContext context, ReadHolder currentReadHolder, Map<Integer, ExcelContentProperty> contentPropertyMap, int propertyIdx, Map<String, Object> fieldMap, Integer totalRow) {
        ExcelContentProperty excelContentProperty = contentPropertyMap.get(propertyIdx);
        ExcelMerge excelMerge = (ExcelMerge) Arrays.stream(excelContentProperty.getField().getDeclaredAnnotations()).filter(item -> item.annotationType().equals(ExcelMerge.class)).findAny().orElse(null);
        ExcelSort excelSort = (ExcelSort) Arrays.stream(excelContentProperty.getField().getDeclaredAnnotations()).filter(item -> item.annotationType().equals(ExcelSort.class)).findAny().orElse(null);
        if (Objects.nonNull(excelMerge)) {
            Integer totalRowNumber = context.readSheetHolder().getApproximateTotalRowNumber();
            if (Objects.nonNull(totalRowNumber) && Objects.isNull(totalRow)) {
                totalRow = totalRowNumber;
            }
            if (excelMerge.rowEnd() > 0) {
                totalRow = excelMerge.rowEnd();
            }
            switch (excelMerge.type()) {
                case COL:
                    List<Object> colList = new ArrayList<>();
                    for (int i = excelMerge.colStart(); i < excelMerge.colEnd(); i++) {
                        CellData data = cellDataMap.get(i);
                        Object value = Objects.isNull(data) || data.getType() == CellDataTypeEnum.EMPTY ? null : ConverterUtils.convertToJavaObject(data, excelContentProperty.getField(),
                                excelContentProperty, currentReadHolder.converterMap(), currentReadHolder.globalConfiguration(),
                                context.readRowHolder().getRowIndex(), i);
                        colList.add(value);
                    }
                    fieldMap.put(excelContentProperty.getField().getName(), colList);
                    break;
                case ROW:
                    cache.put(context.readRowHolder().getRowIndex(), cellDataMap);
                    if (totalRow == context.readRowHolder().getRowIndex() + 1) {
                        List<Object> rowList = new ArrayList<>();
                        for (int row = excelMerge.rowStart(); row < totalRow; row++) {
                            Map<Integer, CellData> rowCell = cache.get(row);
                            CellData data = rowCell.get(propertyIdx);
                            Object value = Objects.isNull(data) || data.getType() == CellDataTypeEnum.EMPTY ? null : ConverterUtils.convertToJavaObject(data, excelContentProperty.getField(),
                                    excelContentProperty, currentReadHolder.converterMap(), currentReadHolder.globalConfiguration(),
                                    context.readRowHolder().getRowIndex(), row);
                            if (Objects.nonNull(excelSort) && Objects.nonNull(rowCell.get(excelSort.colIdx()))) {
                                String sortKey = rowCell.get(excelSort.colIdx()).toString();
                                value = new ExcelSortModel(sortKey, value);
                            }
                            rowList.add(value);
                        }
                        fieldMap.put(excelContentProperty.getField().getName(), rowList);
                    }
                    break;
                case COL_ROW:
                    cache.put(context.readRowHolder().getRowIndex(), cellDataMap);
                    if (totalRow == context.readRowHolder().getRowIndex() + 1) {
                        List<Object> rowList = new ArrayList<>();
                        for (int row = excelMerge.rowStart(); row < totalRow; row++) {
                            for (int col = excelMerge.colStart(); col < excelMerge.colEnd(); col++) {
                                Map<Integer, CellData> rowCell = cache.get(row);
                                CellData data = rowCell.get(col);
                                Object value = Objects.isNull(data) || data.getType() == CellDataTypeEnum.EMPTY ? null : ConverterUtils.convertToJavaObject(data, excelContentProperty.getField(),
                                        excelContentProperty, currentReadHolder.converterMap(), currentReadHolder.globalConfiguration(),
                                        context.readRowHolder().getRowIndex(), row);
                                rowList.add(value);
                            }
                        }
                        fieldMap.put(excelContentProperty.getField().getName(), rowList);
                    }
                    break;
                case ROW_COL:
                    cache.put(context.readRowHolder().getRowIndex(), cellDataMap);
                    if (totalRow == context.readRowHolder().getRowIndex() + 1) {
                        List<Object> rowList = new ArrayList<>();
                        for (int col = excelMerge.colStart(); col < excelMerge.colEnd(); col++) {
                            for (int row = excelMerge.rowStart(); row < totalRow; row++) {
                                Map<Integer, CellData> rowCell = cache.get(row);
                                CellData data = rowCell.get(col);
                                Object value = Objects.isNull(data) || data.getType() == CellDataTypeEnum.EMPTY ? null : ConverterUtils.convertToJavaObject(data, excelContentProperty.getField(),
                                        excelContentProperty, currentReadHolder.converterMap(), currentReadHolder.globalConfiguration(),
                                        context.readRowHolder().getRowIndex(), row);
                                rowList.add(value);
                            }
                        }
                        fieldMap.put(excelContentProperty.getField().getName(), rowList);
                    }
                    break;
                default:
                    break;
            }
        } else {
            Object value = Objects.isNull(cellDataMap.get(propertyIdx)) || cellDataMap.get(propertyIdx).getType() == CellDataTypeEnum.EMPTY ? null :
                    ConverterUtils.convertToJavaObject(cellDataMap.get(propertyIdx), excelContentProperty.getField(),
                            excelContentProperty, currentReadHolder.converterMap(), currentReadHolder.globalConfiguration(),
                            context.readRowHolder().getRowIndex(), propertyIdx);
            if (value != null) {
                fieldMap.put(excelContentProperty.getField().getName(), value);
            }
        }
        return totalRow;
    }

    private Object buildStringList(Map<Integer, CellData> cellDataMap, ReadHolder currentReadHolder,
                                   AnalysisContext context) {
        int index = 0;
        if (context.readWorkbookHolder().getDefaultReturnMap()) {
            Map<Integer, String> map = new LinkedHashMap<Integer, String>(cellDataMap.size() * 4 / 3 + 1);
            for (Map.Entry<Integer, CellData> entry : cellDataMap.entrySet()) {
                Integer key = entry.getKey();
                CellData cellData = entry.getValue();
                while (index < key) {
                    map.put(index, null);
                    index++;
                }
                index++;
                if (cellData.getType() == CellDataTypeEnum.EMPTY) {
                    map.put(key, null);
                    continue;
                }
                map.put(key,
                        (String) ConverterUtils.convertToJavaObject(cellData, null, null, currentReadHolder.converterMap(),
                                currentReadHolder.globalConfiguration(), context.readRowHolder().getRowIndex(), key));
            }
            int headSize = currentReadHolder.excelReadHeadProperty().getHeadMap().size();
            while (index < headSize) {
                map.put(index, null);
                index++;
            }
            return map;
        } else {
            // Compatible with the old code the old code returns a list
            List<String> list = new ArrayList<String>();
            for (Map.Entry<Integer, CellData> entry : cellDataMap.entrySet()) {
                Integer key = entry.getKey();
                CellData cellData = entry.getValue();
                while (index < key) {
                    list.add(null);
                    index++;
                }
                index++;
                if (cellData.getType() == CellDataTypeEnum.EMPTY) {
                    list.add(null);
                    continue;
                }
                list.add(
                        (String) ConverterUtils.convertToJavaObject(cellData, null, null, currentReadHolder.converterMap(),
                                currentReadHolder.globalConfiguration(), context.readRowHolder().getRowIndex(), key));
            }
            int headSize = currentReadHolder.excelReadHeadProperty().getHeadMap().size();
            while (index < headSize) {
                list.add(null);
                index++;
            }
            return list;
        }
    }

}
