package com.alibaba.excel.analysis;

import java.util.ArrayList;
import java.util.List;

import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.model.ModelBuildEventListener;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.support.ExcelTypeEnum;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Created by jipengfei on 17/2/18.
 */
public class AnalyserImpl implements Analyser {

    private AnalysisContext analysisContext;

    public AnalyserImpl(AnalysisContext analysisContext) {
        this.analysisContext = analysisContext;
    }

    public void execute() {

        SaxAnalyser saxAnalyser = getSaxAnalyser();
        saxAnalyser.appendLister(new ModelBuildEventListener());
        saxAnalyser.execute();
        analysisContext.getEventListener().doAfterAllAnalysed(analysisContext);

    }

    private SaxAnalyser getSaxAnalyser() {
        SaxAnalyser saxAnalyser;
        if (ExcelTypeEnum.XLS.equals(analysisContext.getExcelType())) {
            saxAnalyser = new XlsSaxAnalyser(analysisContext);
        } else {
            saxAnalyser = new XlsxSaxAnalyser(analysisContext);
        }
        return saxAnalyser;
    }

    public List<Sheet> getSheets() {
        SaxAnalyser saxAnalyser = getSaxAnalyser();
        return saxAnalyser.getSheets();
    }

}
