package org.nice.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

/**
 * <p>
 * 排序
 * </p>
 *
 * @author WECENG
 * @since 2021/10/15 16:00
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelSortModel implements Serializable {

    /**
     * 排序key
     */
    private String sortKey;

    /**
     * 排序value
     */
    private Object sortValue;

}
