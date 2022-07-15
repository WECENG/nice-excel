package org.nice.excel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.nice.excel.annotation.ExcelMerge;
import org.nice.excel.converter.BigDecimalStringNumberConverter;
import org.nice.excel.enums.MergeEnum;

import java.math.BigDecimal;
import java.util.List;

/**
 * <p>
 * 按列合并excel
 * </p>
 *
 * @author WECENG
 * @since 2021/7/20 10:37
 */
@Data
@NoArgsConstructor
public class ColMergeExcel {

    private String name;

    @ExcelProperty(converter = BigDecimalStringNumberConverter.class)
    @ExcelMerge(type = MergeEnum.COL, colStart = 1, colEnd = 97)
    private List<BigDecimal> dataList;

}
