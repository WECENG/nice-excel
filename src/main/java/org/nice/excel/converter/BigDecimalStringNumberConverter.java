package org.nice.excel.converter;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.alibaba.excel.util.NumberUtils;
import org.springframework.util.StringUtils;

import java.math.BigDecimal;
import java.util.Objects;


/**
 * <p>
 * BigDecimal and String or Number convertor
 * </p>
 *
 * @author WECENG
 * @since 2021/7/5 14:08
 */
public class BigDecimalStringNumberConverter implements Converter<BigDecimal> {

    @Override
    public Class supportJavaTypeKey() {
        return BigDecimal.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.NUMBER;
    }

    @Override
    public BigDecimal convertToJavaData(CellData cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        if (Objects.nonNull(cellData.getNumberValue())) {
            return cellData.getNumberValue();
        }
        if (StringUtils.hasText(cellData.getStringValue())) {
            return NumberUtils.parseBigDecimal(cellData.getStringValue(), contentProperty);
        }
        return null;
    }

    @Override
    public CellData convertToExcelData(BigDecimal value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        return new CellData(value);
    }
}
