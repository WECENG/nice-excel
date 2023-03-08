package org.nice.excel.annotation;


import org.nice.excel.enums.MergeEnum;

import java.lang.annotation.*;

/**
 * <p>
 * 合并
 * </p>
 *
 * @author WECENG
 * @since 2021/7/1 9:31
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelMerge {

    MergeEnum type() default MergeEnum.ROW;

    /**
     * 横向开始
     */
    int rowStart() default -1;

    /**
     * 横向结束
     */
    int rowEnd() default -1;

    /**
     * 纵向开始
     */
    int colStart() default -1;

    /**
     * 纵向结束
     */
    int colEnd() default -1;

    /**
     * 纵向分组
     */
    boolean colGroup() default false;


}
