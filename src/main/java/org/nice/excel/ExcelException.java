package org.nice.excel;

/**
 * <p>
 * excel解析异常类
 * </p>
 *
 * @author WECENG
 * @since 2022/7/14 16:52
 */
public class ExcelException extends Exception {

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(Throwable cause) {
        super(cause);
    }
}
