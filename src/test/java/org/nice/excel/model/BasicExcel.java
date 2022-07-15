package org.nice.excel.model;

import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;

/**
 * <p>
 * 基础excel模板
 * </p>
 *
 * @author WECENG
 * @since 2021/7/5 16:08
 */
@Data
@NoArgsConstructor
public class BasicExcel {

    //    @ExcelProperty(index = 0,converter = AutoConverter.class)   默认从第一列开始
    private String str;

    private BigDecimal number;

}
