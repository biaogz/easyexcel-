package com.alibaba.excel.support;

/**
 * Created by jipengfei on 17/2/18.
 */
public enum  ExcelTypeEnum {
    XLS(".xls"),
    XLSX(".xlsx");
//    CSV(".csv");
    private String value;

    private ExcelTypeEnum(String value){
        this.setValue(value);
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }
}
