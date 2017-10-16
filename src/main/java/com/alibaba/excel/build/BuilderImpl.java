package com.alibaba.excel.build;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.List;

import com.alibaba.excel.context.GenerateContext;
import com.alibaba.excel.metadata.ExcelHead;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Created by jipengfei on 17/2/19.
 */
public class BuilderImpl implements Builder {

    private GenerateContext context;

    public void init(GenerateContext context) {
        this.context = context;

    }

    public void createHead() {
        if (context.getHead() != null && context.getHead().size() > 0) {
            ExcelHead head = new ExcelHead(context.getHead());
            List<ExcelHead.CellRangeModel> list = head.getCellRangeModels();
            int n = context.getCurrentSheet().getLastRowNum();
            if (n > 0) {
                n = n + 4;
            }
            for (ExcelHead.CellRangeModel cellRangeModel : list) {
                CellRangeAddress cra = new CellRangeAddress(cellRangeModel.getFirstRow() + n,
                    cellRangeModel.getLastRow() + n,
                    cellRangeModel.getFirstCol(), cellRangeModel.getLastCol());
                context.getCurrentSheet().addMergedRegion(cra);
            }
            int i = n;
            for (; i < head.getHeadRowNum() + n; i++) {
                createOneRow(context.getCurrentSheet(), head.getHeadByRowNum(i-n), i);
            }
        }
    }

    private void createOneRow(Sheet sheet, List<String> headByRowNum, int rowNum) {
        Row row = sheet.createRow(rowNum);

        //row.setHeightInPoints(30);
        int i = 0;
        for (String excelCell : headByRowNum) {
            createCell(row, excelCell, i, context.getCurrentHeadCellStyle());
            i++;
        }
    }

    private void createCell(Row row, String excelCell, int column, CellStyle cellStyle) {
        Cell cell = row.createCell(column);
        if (excelCell != null) {
            cell.setCellStyle(cellStyle);
            cell.setCellValue(excelCell);
        }
    }

    public void execute() {
        List<Object> data = context.getCurrentData();
        if (data != null && data.size() > 0) {
            int rowNum = context.getCurrentSheet().getLastRowNum();
            for (int i = 0; i < data.size(); i++) {
                int n = i + rowNum + 1;
                Row row = context.getCurrentSheet().createRow(n);
                //  row.setHeightInPoints(30);
                createDtaRow(data.get(i), row);
            }
        }

    }

    private void createDtaRow(Object o, Row row) {
        if (o instanceof List) {
            List<String> stringList = (List<String>)o;
            if (stringList != null && stringList.size() > 0) {

                for (int i = 0; i < stringList.size(); i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellStyle(context.getCurrentContentStyle());
                    cell.setCellValue(stringList.get(i));
                }
            }
        } else {
            Class<?> clazz = context.getCurrentClazz();
            Field[] fields = clazz.getDeclaredFields();
            for (int i = 0; i < fields.length; i++) {
                Field f = fields[i];
                Cell cell = row.createCell(i);
                cell.setCellStyle(context.getCurrentContentStyle());
                String cellValue = null;
                try {
                    cellValue = BeanUtils.getProperty(o, f.getName());
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    e.printStackTrace();
                } catch (NoSuchMethodException e) {
                    e.printStackTrace();
                }
                if (cellValue != null) {
                    cell.setCellValue(cellValue);
                } else {
                    cell.setCellValue("");
                }
            }
        }
    }

}
