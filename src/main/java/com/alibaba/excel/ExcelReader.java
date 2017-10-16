package com.alibaba.excel;

import java.util.List;

import com.alibaba.excel.analysis.Analyser;
import com.alibaba.excel.analysis.AnalyserImpl;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.context.AnalysisContextImpl;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.parameter.AnalysisParam;

/**
 * Excel解析
 * Created by jipengfei on 17/2/18.
 */
public class ExcelReader {

    /**
     * analysisContext
     */
    private AnalysisContext analysisContext;

    /**
     * analyser
     */
    private Analyser analyser;

    public ExcelReader(AnalysisParam param, AnalysisEventListener eventListener) {

        if (param == null) {
            throw new IllegalArgumentException("AnalysisParam  can not null");
        } else if (eventListener == null) {
            throw new IllegalArgumentException("AnalysisEventListener can not null");
        } else if (param.getIn() == null) {
            throw new IllegalArgumentException("InputStream can not null");
        }

        analysisContext = new AnalysisContextImpl(param.getIn(), param.getExcelTypeEnum(), param.getCustomContent(),
            eventListener);
        analyser = new AnalyserImpl(analysisContext);
    }

    /**
     * 读一个sheet，且没有模型映射
     */
    public void read() {
        analyser.execute();
    }

    /**
     * 读某个sheet，且没有模型映射
     *
     * @param sheetParam
     */
    public void read(Sheet sheetParam) {
        analysisContext.setCurrentSheet(sheetParam);
        analyser.execute();
    }

    /**
     * 读某个sheet，且有模型映射
     *
     * @param sheetParam
     * @param clazz
     */
    public void read(Sheet sheetParam, Class<? extends BaseRowModel> clazz) {
        analysisContext.setCurrentSheet(sheetParam);
        analysisContext.setCurrentClazz(clazz);
        analyser.execute();
    }

    /**
     * 读取excel中有哪些sheet
     * @return
     */
    public List<Sheet> getSheets() {
        return analyser.getSheets();
    }

}
