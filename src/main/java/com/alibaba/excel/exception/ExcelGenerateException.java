package com.alibaba.excel.exception;

/**
 * Created by jipengfei on 17/2/18.
 */
public class ExcelGenerateException extends RuntimeException {

    public ExcelGenerateException() {
    }

    public ExcelGenerateException(String message) {
        super(message);
    }

    public ExcelGenerateException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelGenerateException(Throwable cause) {
        super(cause);
    }
}
