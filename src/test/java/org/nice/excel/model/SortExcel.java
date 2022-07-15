package org.nice.excel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.nice.excel.annotation.ExcelMerge;
import org.nice.excel.annotation.ExcelSort;
import org.nice.excel.converter.BigDecimalStringNumberConverter;
import org.nice.excel.enums.MergeEnum;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;

/**
 * <p>
 * 按行合并excel
 * </p>
 *
 * @author WECENG
 * @since 2021/7/5 10:28
 */
@Data
@NoArgsConstructor
public class SortExcel {

    @DateTimeFormat("yyyy-MM-dd")
    @ExcelProperty(index = 0)
    private Date date;

    @ExcelMerge(type = MergeEnum.ROW, rowStart = 1, rowEnd = 97)
    @ExcelProperty(index = 2, converter = BigDecimalStringNumberConverter.class)
    @ExcelSort(colIdx = 1)
    private List<BigDecimal> dataList;

}
