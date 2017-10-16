package com.alibaba.excel.event;

import com.alibaba.excel.context.AnalysisContext;

/**
 * Created by jipengfei on 17/2/18.
 */
public abstract class AnalysisEventListener {

    /**
     *  when analysis one row trigger invoke function
     * @param object one row data
     * @param context analysis context
     */
    public abstract void invoke(Object object,AnalysisContext context);

    /**
     *
     * if have something to do after all  analysis
     * @param context
     */
    public abstract void doAfterAllAnalysed(AnalysisContext context);
}
