package org.nice.excel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.nice.excel.annotation.ExcelGroup;
import org.nice.excel.annotation.ExcelMerge;
import org.nice.excel.converter.BigDecimalStringNumberConverter;
import org.nice.excel.enums.MergeEnum;

import java.math.BigDecimal;
import java.util.List;

/**
 * <p>
 * 分组后合并
 * </p>
 *
 * @author WECENG
 * @since 2023/3/7 09:58
 */
@Data
@NoArgsConstructor
@ExcelGroup(fields = {"group"}, fill = true)
public class ColGroupMerge {

    @ExcelProperty(index = 0)
    @ExcelMerge(type = MergeEnum.ROW, rowStart = 1)
    private String group;

    @ExcelProperty(index = 2, converter = BigDecimalStringNumberConverter.class)
    @ExcelMerge(type = MergeEnum.COL, colStart = 2, colEnd = 26, colGroup = true)
    private List<BigDecimal> dataList;

}
