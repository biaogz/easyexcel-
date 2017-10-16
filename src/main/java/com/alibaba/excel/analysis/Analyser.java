package com.alibaba.excel.analysis;

import java.util.List;

import com.alibaba.excel.metadata.Sheet;

/**
 * Created by jipengfei on 17/2/18.
 */
public interface Analyser {

    void execute();

    List<Sheet> getSheets();
}
