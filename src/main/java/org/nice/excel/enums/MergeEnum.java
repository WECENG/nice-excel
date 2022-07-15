package org.nice.excel.enums;

/**
 * <p>
 * 合并方式枚举
 * </p>
 *
 * @author WECENG
 * @since 2021/7/1 9:31
 */
public enum MergeEnum {
    /**
     * 横向合并
     */
    ROW("横向合并"),
    /**
     * 纵向合并
     */
    COL("纵向合并"),
    /**
     * 横纵合并
     */
    ROW_COL("横纵合并"),
    /**
     * 纵横合并
     */
    COL_ROW("纵横合并");

    public final String text;

    MergeEnum(String text) {
        this.text = text;
    }

    public String getText() {
        return text;
    }
}
