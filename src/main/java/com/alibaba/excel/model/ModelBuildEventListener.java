package com.alibaba.excel.model;

import java.util.ArrayList;
import java.util.List;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;

/**
 * Created by jipengfei on 17/2/18.
 */
public class ModelBuildEventListener extends AnalysisEventListener {

    private ModelBuild modelBuild;

    public ModelBuildEventListener() {
        modelBuild = new ModelBuild();
    }

    private int rowNum;

    @Override
    public void invoke(Object object, AnalysisContext context) {
        rowNum = context.getCurrentRowNum();
        AnalysisEventListener userListener = context.getEventListener();
        List list = (List)object;
        List obj = new ArrayList();
        obj.addAll(list);
        Sheet sheet = context.getCurrentSheet();
        if (sheet == null) {
            userListener.invoke(obj, context);
            return;
        }
        if (rowNum < sheet.getHeadLineMun()) {
            if (rowNum == sheet.getHeadLineMun() - 1) {
                modelBuild.buildHead(context.getHead(), obj);
            }
        } else if (context.getCurrentClazz() == null) {
            userListener.invoke(obj, context);
        } else {
            Class c = context.getCurrentClazz();
            try {
                Object resultModel = c.newInstance();
                modelBuild.buildModel(resultModel, context.getCurrentClazz(), obj, context.getHead());
                userListener.invoke(resultModel, context);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }


    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

    }
}
