package com.alibaba.excel.event;

import java.util.List;

/**
 * Created by jipengfei on 17/2/18.
 */
public interface AnalysisEventRegisterCenter {

    /**
     *
     * @param listener
     */
    void appendLister(AnalysisEventListener listener);

    /**
     *
     * @return
     */
    List<AnalysisEventListener> getAllRegister();


    void notifyListeners();
}
