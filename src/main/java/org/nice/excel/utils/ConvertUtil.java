package org.nice.excel.utils;

import org.springframework.util.CollectionUtils;

import java.util.*;

/**
 * <p>
 * 转换器工具类
 * </p>
 *
 * @author WECENG
 * @since 2021/5/19 17:03
 */
public class ConvertUtil {

    /**
     * 根据udx收集下标
     *
     * @param udxMap
     * @param udx
     * @return
     */
    public static List<Integer> collectDataByUdx(Map<Integer, Object> udxMap, Object udx) {
        List<Integer> idxList = new ArrayList<>();
        if (CollectionUtils.isEmpty(udxMap) || Objects.isNull(udx)) {
            return idxList;
        }
        udxMap.forEach((idx, val) -> {
            if (udx.equals(val)) {
                idxList.add(idx);
            }
        });
        return idxList;
    }


    /**
     * 根据下标集合匹配数据
     *
     * @param idxList 下标集合
     * @param allData 数据集
     * @return
     */
    public static List<Object> matchData(List<Integer> idxList, List<Object> allData) {
        List<Object> dataList = new ArrayList<>();
        int size = allData.stream()
                .filter(item -> item instanceof List)
                .map(item -> (List<?>) item)
                .map(List::size).min(Integer::compareTo)
                .orElse(1);
        for (int i = 0; i < size; i++) {
            for (Integer idx : idxList) {
                Object val = allData.get(idx);
                if (val instanceof List) {
                    List<?> rowList = (List<?>) val;
                    dataList.add(rowList.get(i));
                } else {
                    dataList.add(val);
                }
            }
        }
        return dataList;
    }


    /**
     * 根据下标集合匹配数据
     *
     * @param idxList 下标集合
     * @param allData 数据集
     * @return
     */
    public static Map<Integer, Object> matchData(List<Integer> idxList, Map<Integer, Object> allData) {
        Map<Integer, Object> dataMap = new HashMap<>(8);
        idxList.forEach(idx -> dataMap.put(idx, allData.get(idx)));
        return dataMap;
    }

    /**
     * 逐item拼接
     *
     * @param list1
     * @param list2
     * @return
     */
    public static List<String> listItemComb(List<String> list1, List<String> list2) {
        List<String> resultList = new ArrayList<>();
        if (CollectionUtils.isEmpty(list1) || CollectionUtils.isEmpty(list2)) {
            return resultList;
        }
        for (int i = 0; i < list1.size() && i < list2.size(); i++) {
            resultList.add(list1.get(i) + list2.get(i));
        }
        return resultList;
    }


    /**
     * 用前一个element填充为null的element
     *
     * @param list 待处理集合
     */
    public static void fillPreNullable(List<Object> list) {
        ListIterator<Object> iterator = list.listIterator();
        Object prev = null;
        while (iterator.hasNext()) {
            Object current = iterator.next();
            if (current == null) {
                iterator.set(prev);
            } else {
                prev = current;
            }
        }
    }

}
