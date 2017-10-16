package com.alibaba.excel.analysis;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.event.AnalysisEventRegisterCenter;
import com.alibaba.excel.metadata.Sheet;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by jipengfei on 17/2/18.
 */
public abstract class SaxAnalyser implements AnalysisEventRegisterCenter, Analyser {


    private List<AnalysisEventListener> registers = new ArrayList<AnalysisEventListener>();

    public void appendLister(AnalysisEventListener listener) {
        registers.add(listener);
    }

    public List<AnalysisEventListener> getAllRegister() {
        return registers;
    }

    public List<Sheet> getSheets() {
        return null;
    }
}
