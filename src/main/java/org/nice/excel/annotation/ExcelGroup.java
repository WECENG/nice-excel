package org.nice.excel.annotation;

import java.lang.annotation.*;

/**
 * <p>
 * 分组
 * </p>
 *
 * @author WECENG
 * @since 2021/7/1 9:31
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelGroup {
    /**
     * 分组属性
     */
    String[] fields();

    /**
     * 填充单元各合并
     */
    boolean fill() default false;
}
