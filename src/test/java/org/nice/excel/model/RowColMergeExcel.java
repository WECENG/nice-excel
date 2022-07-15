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
 * 行列合并excel
 * </p>
 *
 * @author WECENG
 * @since 2021/7/5 16:32
 */
@Data
@NoArgsConstructor
public class RowColMergeExcel {

    @ExcelProperty(converter = BigDecimalStringNumberConverter.class)
    @ExcelMerge(type = MergeEnum.ROW_COL, rowStart = 1, rowEnd = 5, colStart = 1, colEnd = 25)
    private List<BigDecimal> dataList;

}
