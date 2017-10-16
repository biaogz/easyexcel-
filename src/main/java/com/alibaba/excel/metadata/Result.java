package com.alibaba.excel.metadata;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by jipengfei on 17/2/18.
 */
public class Result<T extends BaseRowModel>{

    private List<T> datas = new ArrayList<T>();

    private List<List<String>> head = new ArrayList<List<String>>();

}
