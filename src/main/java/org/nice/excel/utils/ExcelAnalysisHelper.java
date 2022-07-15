package org.nice.excel.utils;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.CellExtra;

import java.lang.reflect.Field;
import java.util.List;

/**
 * 单元格合并helper类
 *
 * @author chenwc
 * @date 2021/07/09 9:44
 */
public class ExcelAnalysisHelper<T> {

    /**
     * 处理合并单元格
     *
     * @param data               解析数据
     * @param extraMergeInfoList 合并单元格信息
     * @param columnNames        标题名称
     * @param headRowNumber      表头行
     * @return 填充好的解析数据
     */
    public List<T> explainMergeData(List<T> data, List<CellExtra> extraMergeInfoList, List<String> columnNames, int headRowNumber) {
        //循环所有合并单元格信息
        extraMergeInfoList.forEach(cellExtra -> {
            int firstRowIndex = cellExtra.getFirstRowIndex() - headRowNumber;
            int lastRowIndex = cellExtra.getLastRowIndex() - headRowNumber;
            int firstColumnIndex = cellExtra.getFirstColumnIndex();
            int lastColumnIndex = cellExtra.getLastColumnIndex();
            //获取初始值
            if (firstRowIndex < data.size()) {
                Object initValue = getInitValueFromList(firstRowIndex, firstColumnIndex, data, columnNames);
                //设置值
                for (int i = firstRowIndex; i <= lastRowIndex; i++) {
                    for (int j = firstColumnIndex; j <= lastColumnIndex; j++) {
                        setInitValueToList(initValue, i, j, data, columnNames);
                    }
                }
            }
        });
        return data;
    }

    /**
     * 设置合并单元格的值
     *
     * @param filedValue  值
     * @param rowIndex    行
     * @param columnIndex 列
     * @param data        解析数据
     */
    public void setInitValueToList(Object filedValue, Integer rowIndex, Integer columnIndex, List<T> data, List<String> columnNames) {
        T object = data.get(rowIndex);
        for (Field field : object.getClass().getDeclaredFields()) {
            field.setAccessible(true);
            ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
            if (annotation != null) {
                StringBuilder names = new StringBuilder();
                for (String s : annotation.value()) {
                    names.append(s);
                }
                int index = columnNames.indexOf(names.toString());
                if (index == columnIndex) {
                    try {
                        field.set(object, filedValue);
                        break;
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }
            }
        }
    }


    /**
     * 获取合并单元格的初始值
     * rowIndex对应list的索引
     * columnIndex对应实体内的字段
     *
     * @param firstRowIndex    起始行
     * @param firstColumnIndex 起始列
     * @param data             列数据
     * @param nameList         excel表头名称
     * @return 初始值
     */
    private Object getInitValueFromList(Integer firstRowIndex, Integer firstColumnIndex, List<T> data, List<String> nameList) {
        Object filedValue = null;
        T object = data.get(firstRowIndex);
        Field[] fields = object.getClass().getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true);
            ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
            if (annotation != null) {
                StringBuilder names = new StringBuilder();
                for (String s : annotation.value()) {
                    names.append(s);
                }
                int index = nameList.indexOf(names.toString());
                if (index == firstColumnIndex) {
                    try {
                        filedValue = field.get(object);
                        break;
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }

                }
            }
        }
        return filedValue;
    }

}