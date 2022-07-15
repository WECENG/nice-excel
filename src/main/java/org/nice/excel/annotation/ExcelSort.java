package org.nice.excel.annotation;

import java.lang.annotation.*;

/**
 * <p>
 * 排序
 * </p>
 *
 * @author WECENG
 * @since 2021/10/15 14:56
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelSort {

    /**
     * 排序列索引
     *
     * @return 列索引
     */
    int colIdx();

}
