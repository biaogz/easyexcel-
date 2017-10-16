package com.alibaba.excel.context;

import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by jipengfei on 17/2/18.
 */
public class AnalysisContextImpl implements AnalysisContext {

    private Object custom;

    private Sheet currentSheet;

    private ExcelTypeEnum excelType;

    private InputStream inputStream;

    private AnalysisEventListener eventListener;

    private List<List<String>> head;

    private Class<?> currentClazz;

    private Integer currentRownNum;

    private Integer totalCount;


    public AnalysisContextImpl(InputStream inputStream,ExcelTypeEnum excelTypeEnum,Object custom, AnalysisEventListener listener){
        this.custom = custom;
        this.eventListener = listener;
        this.inputStream = inputStream;
        this.excelType = excelTypeEnum;
        this.head = new ArrayList<List<String>>();
    }



    public void setCurrentSheet(Sheet currentSheet) {
        this.currentSheet = currentSheet;
        this.currentClazz = currentSheet.getClazz();
    }

    public ExcelTypeEnum getExcelType() {
        return excelType;
    }

    public void setExcelType(ExcelTypeEnum excelType) {
        this.excelType = excelType;
    }

    public Object getCustom() {
        return custom;
    }

    public void setCustom(Object custom) {
        this.custom = custom;
    }

    public Sheet getCurrentSheet() {
        return currentSheet;
    }


    public InputStream getInputStream() {
        return inputStream;
    }

    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }

    public AnalysisEventListener getEventListener() {
        return eventListener;
    }

    public void setEventListener(AnalysisEventListener eventListener) {
        this.eventListener = eventListener;
    }

    public List<List<String>> getHead() {
        return head;
    }

    public Integer getCurrentRowNum() {
        return this.currentRownNum;
    }

    public void setCurrentRownNum(Integer row) {
            this.currentRownNum = row;
    }

    public void setHead(List<List<String>> head) {
        this.head = head;
    }

    public Class<?> getCurrentClazz() {
        return currentClazz;
    }

    public void setCurrentClazz(Class<?> currentClazz) {
        this.currentClazz = currentClazz;
    }

    public Integer getTotalCount() {
        return totalCount;
    }

    public void setTotalCount(Integer totalCount) {
        this.totalCount = totalCount;
    }
}
