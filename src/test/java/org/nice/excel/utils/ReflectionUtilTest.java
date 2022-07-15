package org.nice.excel.utils;


import org.junit.jupiter.api.Test;
import org.nice.excel.listener.GroupByExcelListener;
import org.nice.excel.listener.RowMergeExcelListener;
import org.nice.excel.model.GroupByExcel;
import org.nice.excel.model.RowMergeExcel;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class ReflectionUtilTest {

    @Test
    public void getGenericsClass() {
        Class<?> sunPowerClazz = ReflectionUtil.getGenericsClass(RowMergeExcelListener.class);
        assertEquals(sunPowerClazz, RowMergeExcel.class);
        Class<?> bootOffClazz = ReflectionUtil.getGenericsClass(GroupByExcelListener.class);
        assertEquals(bootOffClazz, GroupByExcel.class);
    }

}