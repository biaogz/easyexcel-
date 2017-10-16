package com.alibaba.excel;

import com.alibaba.excel.build.Builder;
import com.alibaba.excel.build.BuilderImpl;
import com.alibaba.excel.context.GenerateContext;
import com.alibaba.excel.context.GenerateContextImpl;
import com.alibaba.excel.exception.ExcelGenerateException;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.parameter.ExcelWriteParam;
import com.alibaba.excel.parameter.GenerateParam;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * 生成excel
 * Created by jipengfei on 17/2/18.
 */
public class ExcelWriter {

    private GenerateContext context;

    private Builder builder = new BuilderImpl();

    /**
     * 由于不能生成多sheet，故废弃掉
     * @param generateParam
     */
    @Deprecated
    public ExcelWriter(GenerateParam generateParam) {
        context = new GenerateContextImpl(generateParam.getOutputStream(), generateParam.getType());
        Sheet sheet = new Sheet(1, 0, generateParam.getClazz(), generateParam.getSheetName(), null);
        builder.init(context);
        context.buildCurrentSheet(sheet, builder);
    }

    public ExcelWriter(ExcelWriteParam writeParam) {
        context = new GenerateContextImpl(writeParam.getOutputStream(), writeParam.getType());
        builder.init(context);
    }

    /**
     * 由于不能生成多sheet，故废弃掉
     * @param data
     * @return
     */
    @Deprecated
    public ExcelWriter write(List data) {
        context.setCurrentData(data);
        builder.execute();
        return this;
    }

    /**
     * 可生成多sheet,每个sheet一张表
     *
     * @param data       List<Object> Object type is <? extends BaseRowModel> or List<String>
     * @param sheetParam
     * @return
     */
    public ExcelWriter write(List data, Sheet sheetParam) {
        context.buildCurrentSheet(sheetParam, builder);
        context.setCurrentData(data);
        builder.execute();
        return this;
    }

    /**
     * 可生成多sheet,每个sheet一张表
     *
     * @param data       List<Object> Object type is <? extends BaseRowModel> or List<String>
     * @param sheetParam
     * @return
     */
    public ExcelWriter write(List data, Sheet sheetParam, Table table) {
        context.buildCurrentSheet(sheetParam, builder);
        context.buildTable(table,builder);
        context.setCurrentData(data);
        builder.execute();
        return this;
    }

    public void finish() {
        try {
            context.getWorkbook().write(context.getOutputStream());
        } catch (IOException e) {
            throw new ExcelGenerateException(e);
        }
    }
}
