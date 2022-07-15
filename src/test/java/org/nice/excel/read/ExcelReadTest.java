package org.nice.excel.read;

import com.alibaba.excel.EasyExcel;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.nice.excel.listener.*;
import org.nice.excel.model.*;

import java.io.InputStream;

/**
 * <p>
 * excel单元测试类
 * </p>
 *
 * @author WECENG
 * @since 2021/7/5 10:23
 */
public class ExcelReadTest {

    @Test
    @DisplayName("基本excel")
    public void basicExcelTest() {
        InputStream is = this.getClass().getClassLoader().getResourceAsStream("dir/basicExcel.xlsx");
        EasyExcel.read(is, BasicExcel.class, new CustomModelEventListener())
                .useDefaultListener(false)
                .registerReadListener(new BasicExcelListener())
                .sheet()
                .doRead();
    }

    @Test
    @DisplayName("按行合并excel")
    public void rowMergeTest() {
        InputStream is = this.getClass().getClassLoader().getResourceAsStream("dir/rowMerge.xlsx");
        EasyExcel.read(is, new CustomModelEventListener())
                .useDefaultListener(false)
                .registerReadListener(new RowMergeExcelListener())
                .head(RowMergeExcel.class)
                .doReadAll();
    }

    @Test
    @DisplayName("按列合并excel")
    public void colMergeTest() {
        InputStream is = this.getClass().getClassLoader().getResourceAsStream("dir/colMerge.xls");
        EasyExcel.read(is, new CustomModelEventListener())
                .useDefaultListener(false)
                .registerReadListener(new ColMergeExcelListener())
                .head(ColMergeExcel.class)
                .doReadAll();
    }

    @Test
    @DisplayName("行列双向合并excel")
    public void rowColMergeTest() {
        InputStream is = this.getClass().getClassLoader().getResourceAsStream("dir/rowColMerge.xls");
        EasyExcel.read(is, new CustomModelEventListener())
                .useDefaultListener(false)
                .registerReadListener(new RowColMergeExcelListener())
                .head(RowColMergeExcel.class)
                .doReadAll();
    }

    @Test
    @DisplayName("分组excel")
    public void groupByExcelTest() {
        InputStream is = this.getClass().getClassLoader().getResourceAsStream("dir/groupBy.xls");
        EasyExcel.read(is, GroupByExcel.class, new CustomModelEventListener())
                .useDefaultListener(false)
                .registerReadListener(new GroupByExcelListener())
                .sheet()
                .doRead();
    }

    @Test
    @DisplayName("排序excel")
    public void sortExcelTest() {
        InputStream is = this.getClass().getClassLoader().getResourceAsStream("dir/sort.xlsx");
        EasyExcel.read(is, SortExcel.class, new CustomModelEventListener())
                .useDefaultListener(false)
                .registerReadListener(new SortExcelListener())
                .sheet()
                .doRead();
    }

}
