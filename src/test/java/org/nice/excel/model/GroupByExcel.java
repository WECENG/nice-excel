package org.nice.excel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.converters.string.StringStringConverter;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.nice.excel.annotation.ExcelGroup;
import org.nice.excel.annotation.ExcelMerge;
import org.nice.excel.annotation.ExcelSort;
import org.nice.excel.converter.BigDecimalStringNumberConverter;
import org.nice.excel.enums.MergeEnum;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;

/**
 * <p>
 * 分组excel
 * </p>
 *
 * @author WECENG
 * @since 2021/7/5 14:23
 */
@Data
@NoArgsConstructor
@ExcelGroup(fields = {"group1", "group2", "group3"})
public class GroupByExcel {

    @ExcelMerge(type = MergeEnum.ROW, rowStart = 1)
    @ExcelProperty(index = 0)
    private String group1;

    @ExcelProperty(index = 1)
    @ExcelMerge(type = MergeEnum.ROW, rowStart = 1)
    private String group2;

    @ExcelProperty(index = 2)
    @DateTimeFormat("yyyy-MM-dd")
    @ExcelMerge(type = MergeEnum.ROW, rowStart = 1)
    private Date group3;

    @ExcelProperty(index = 4, converter = BigDecimalStringNumberConverter.class)
    @ExcelMerge(type = MergeEnum.ROW, rowStart = 1)
    private List<BigDecimal> dataList;

}
