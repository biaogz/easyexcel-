package com.alibaba.excel.metadata;

import java.util.List;

/**
 * @author jipengfei
 * @date 2017/05/16
 */
public class Table {
    /**
     * 对用的表头模型
     */
    private Class<? extends BaseRowModel> clazz;

    /**
     * 对用的表头层级树,用于clazz不确定时候，动态生成表头
     */
    private List<List<String>> head;

    private Integer tableNo;

    private TableStyle tableStyle;

    public TableStyle getTableStyle() {
        return tableStyle;
    }

    public void setTableStyle(TableStyle tableStyle) {
        this.tableStyle = tableStyle;
    }

    public Table(Integer tableNo) {
        this.tableNo = tableNo;
    }

    public Class<? extends BaseRowModel> getClazz() {
        return clazz;
    }

    public void setClazz(Class<? extends BaseRowModel> clazz) {
        this.clazz = clazz;
    }

    public List<List<String>> getHead() {
        return head;
    }

    public void setHead(List<List<String>> head) {
        this.head = head;
    }

    public Integer getTableNo() {
        return tableNo;
    }

    public void setTableNo(Integer tableNo) {
        this.tableNo = tableNo;
    }
}
