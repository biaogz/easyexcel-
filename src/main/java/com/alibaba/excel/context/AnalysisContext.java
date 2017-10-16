package com.alibaba.excel.context;

import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;

import java.io.InputStream;
import java.util.List;

/**
 * Created by jipengfei on 17/2/18.
 */
public interface AnalysisContext {


    /**
     *
     * @param clazz
     */
    void setCurrentClazz(Class<?> clazz);


    /**
     *
     * @return
     */
    Class<?> getCurrentClazz();

    /**
     *
     * @return
     */
    Object getCustom();

    /**
     *
     * @return current analysis sheet
     */
    Sheet getCurrentSheet();

    /**
     *
     * @param sheet
     */
    void setCurrentSheet(Sheet sheet);

    /**
     *
     * @return excel type
     */
    ExcelTypeEnum getExcelType();


    /**
     *
     * @return file io
     */
    InputStream getInputStream();


    /**
     *
     * @return
     */
    AnalysisEventListener getEventListener();


    /**
     *
     * @return
     */
    List<List<String>> getHead();

    /**
     *
     * @return
     */
    Integer getCurrentRowNum();

    /**
     *
     */
    void setCurrentRownNum(Integer row);

    /**
     * 返回当前sheet共有多少行数据，仅限07版excel
     *
     * @return
     */
    Integer getTotalCount();

    /**
     *
     * @param totalCount
     */
    void setTotalCount(Integer totalCount) ;
}
