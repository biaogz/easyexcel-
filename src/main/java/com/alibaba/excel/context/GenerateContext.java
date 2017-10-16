package com.alibaba.excel.context;

import java.io.OutputStream;
import java.util.List;

import com.alibaba.excel.build.Builder;
import com.alibaba.excel.metadata.Table;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Created by jipengfei on 17/2/19.
 */
public interface GenerateContext {

    /**
     * @return
     */
    Class<?> getCurrentClazz();

    /**
     * @return current analysis sheet
     */
    Sheet getCurrentSheet();

    /**
     * @return
     */
    List<List<String>> getHead();

    /**
     * @param head
     */
    void setHead(List<List<String>> head);


    /**
     *
     * @return
     */
    CellStyle getCurrentHeadCellStyle();

    /**
     *
     * @return
     */
    CellStyle getCurrentContentStyle();
    /**
     * @param data
     */
    void setCurrentData(List<Object> data);

    /**
     * @return
     */
    List<Object> getCurrentData();

    /**
     * @return
     */
    Workbook getWorkbook();

    /**
     * @return
     */
    OutputStream getOutputStream();

    /**
     *
     * @param sheet
     * @param builder
     */
    void buildCurrentSheet(com.alibaba.excel.metadata.Sheet sheet, Builder builder);

    /**
     *
     * @param table
     * @param builder
     */
    void buildTable(Table table, Builder builder);

}


