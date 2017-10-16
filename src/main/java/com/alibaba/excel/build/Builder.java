package com.alibaba.excel.build;

import com.alibaba.excel.context.GenerateContext;
import com.alibaba.excel.support.ExcelTypeEnum;

/**
 * Created by jipengfei on 17/2/19.
 */
public interface Builder {

    void init(GenerateContext context);

    void execute();

    void createHead();

}
