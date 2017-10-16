package com.alibaba.excel.parameter;

import java.io.OutputStream;

import com.alibaba.excel.support.ExcelTypeEnum;

/**
 * @author jipengfei
 * @date 2017/05/15
 */
public class ExcelWriteParam {

    /**
     * 文件输出流
     */
    private OutputStream outputStream;

    /**
     * Excel类型
     */
    private ExcelTypeEnum type;



    public ExcelWriteParam(OutputStream outputStream, ExcelTypeEnum type) {
        this.outputStream = outputStream;
        this.type = type;

    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public void setOutputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    public ExcelTypeEnum getType() {
        return type;
    }

    public void setType(ExcelTypeEnum type) {
        this.type = type;
    }
}
