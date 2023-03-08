package org.nice.excel;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.IoUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.nice.excel.annotation.ExcelGroup;
import org.nice.excel.annotation.ExcelMerge;
import org.nice.excel.converter.BigDecimalStringNumberConverter;
import org.nice.excel.enums.MergeEnum;
import org.nice.excel.listener.CustomModelEventListener;
import org.nice.excel.listener.TypeMappingListener;
import org.springframework.util.CollectionUtils;

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * <p>
 * 分组统计
 * </p>
 *
 * @author WECENG
 * @since 2022/12/9 13:01
 */
public class GroupStatistics {

    public static void main(String[] args) throws IOException {
        String sourceDir = "/Users/chenweicheng/Documents/rawFiles";
        String targetDir = "/Users/chenweicheng/Documents/files";
        String sheetName = "搜索词统计";
        if (FileUtil.isDirectory(sourceDir)) {
            List<File> files = FileUtil.loopFiles(sourceDir);
            for (File file : files) {
                InputStream is = null;
                FileOutputStream os = null;
                try {
                    is = new FileInputStream(file);
                    List<WriteExcel> writeDataList = new ArrayList<>();
                    File outFile = new File(targetDir + "/" + file.getName());
                    outFile.createNewFile();
                    os = new FileOutputStream(outFile);
                    byte[] bytes = IoUtil.readBytes(is);
                    ByteArrayInputStream inputStream = new ByteArrayInputStream(bytes);
                    IoUtil.copy(inputStream, os);
                    inputStream.reset();
                    EasyExcel.read(inputStream, GroupExcel.class, new CustomModelEventListener())
                            .useDefaultListener(false)
                            .registerReadListener(new TypeMappingListener<List<GroupExcel>>() {
                                @Override
                                public void doInvoke(List<GroupExcel> groupExcels, AnalysisContext context, Map<?, ?> customMap) {
                                    List<WriteExcel> writeExcelList = groupExcels.stream()
                                            .map(item ->
                                                    new WriteExcel(
                                                            item.getKey(),
                                                            GroupStatistics.sumList(item.getNumList()),
                                                            GroupStatistics.avg(GroupStatistics.sumList(item.getSpentList()),GroupStatistics.sumList(item.getNumList())),
                                                            GroupStatistics.sumList(item.getSpentList()),
                                                            GroupStatistics.sumList(item.getFeeList()),
                                                            GroupStatistics.sumList(item.getOrderList()),
                                                            GroupStatistics.sumList(item.getSellList())
                                                    )
                                            )
                                            .collect(Collectors.toList());
                                    writeDataList.addAll(writeExcelList);
                                }

                                @Override
                                public void post(AnalysisContext context, Map<?, ?> customMap) {
                                }
                            })
                            .sheet()
                            .doRead();
                    EasyExcel.write(outFile, WriteExcel.class).withTemplate(file.getPath()).autoCloseStream(true).sheet(sheetName).doWrite(writeDataList);
                } catch (Exception e) {
                    System.out.println(e.getMessage());
                } finally {
                    if (os != null) {
                        os.close();
                    }
                    if (is != null) {
                        is.close();
                    }
                }
            }
        }

    }

    @Data
    @NoArgsConstructor
    @ExcelGroup(fields = {"key"})
    public static class GroupExcel {

        @ExcelProperty(value = "客户搜索词")
        @ExcelMerge(type = MergeEnum.ROW, rowStart = 1)
        private String key;

        @ExcelProperty(value = "点击量", converter = BigDecimalStringNumberConverter.class)
        @ExcelMerge(type = MergeEnum.ROW, rowStart = 1)
        private List<BigDecimal> numList;

        @ExcelProperty(value = "花费", converter = BigDecimalStringNumberConverter.class)
        @ExcelMerge(type = MergeEnum.ROW, rowStart = 1)
        private List<BigDecimal> spentList;

        @ExcelProperty(value = "7天总销售额", converter = BigDecimalStringNumberConverter.class)
        @ExcelMerge(type = MergeEnum.ROW, rowStart = 1)
        private List<BigDecimal> feeList;

        @ExcelProperty(value = "7天总订单数(#)", converter = BigDecimalStringNumberConverter.class)
        @ExcelMerge(type = MergeEnum.ROW, rowStart = 1)
        private List<BigDecimal> orderList;

        @ExcelProperty(value = "7天总销售量(#)", converter = BigDecimalStringNumberConverter.class)
        @ExcelMerge(type = MergeEnum.ROW, rowStart = 1)
        private List<BigDecimal> sellList;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class WriteExcel {

        @ExcelProperty(value = "客户搜索词", index = 0)
        private String key;

        @ExcelProperty(value = "点击量", index = 1)
        private BigDecimal num;

        @ExcelProperty(value = "单次点击成本", index = 2)
        private BigDecimal clickCost;

        @ExcelProperty(value = "花费", index = 3)
        private BigDecimal spent;

        @ExcelProperty(value = "总销售额", index = 4)
        private BigDecimal fee;

        @ExcelProperty(value = "总订单数", index = 5)
        private BigDecimal order;

        @ExcelProperty(value = "总销售量", index = 6)
        private BigDecimal sell;
    }


    public static BigDecimal sumList(List<BigDecimal> numList) {
        if (CollectionUtils.isEmpty(numList)) {
            return BigDecimal.ZERO;
        }
        BigDecimal sum = BigDecimal.ZERO;
        for (BigDecimal item : numList) {
            sum = sum.add(item);
        }
        return sum;
    }

    public static BigDecimal avg(BigDecimal a, BigDecimal b) {
        if (Objects.isNull(a) || Objects.isNull(b)) {
            return BigDecimal.ZERO;
        }
        if (b.compareTo(BigDecimal.ZERO) == 0){
            return BigDecimal.ZERO;
        }
        return a.divide(b,2, RoundingMode.HALF_UP);
    }

}
