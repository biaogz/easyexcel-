package com.alibaba.excel.context;

import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.build.Builder;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.metadata.TableStyle;
import com.alibaba.excel.support.ExcelTypeEnum;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Created by jipengfei on 17/2/19.
 */
public class GenerateContextImpl implements GenerateContext {

    private Class<?> currentClazz;

    private Sheet currentSheet;

    private String currentSheetName;

    private ExcelTypeEnum excelType;

    private List<List<String>> head;

    private List<Object> currentData;

    private Workbook workbook;

    private OutputStream outputStream;

    private Map<Integer, Sheet> sheetMap = new ConcurrentHashMap<Integer, Sheet>();

    private Map<Integer, Table> tableMap = new ConcurrentHashMap<Integer, Table>();

    private CellStyle defaultCellStyle;

    private CellStyle currentHeadCellStyle;

    private CellStyle currentContentCellStyle;

    public GenerateContextImpl(OutputStream out, ExcelTypeEnum excelType) {
        if (ExcelTypeEnum.XLS.equals(excelType)) {
            this.workbook = new HSSFWorkbook();
        } else {
            this.workbook = new XSSFWorkbook();
        }
        this.outputStream = out;
        this.defaultCellStyle = buildDefaultCellStyle();
    }

    private CellStyle buildDefaultCellStyle() {
        CellStyle newCellStyle = this.workbook.createCellStyle();
        Font font = this.workbook.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short)14);
        font.setBold(true);
        newCellStyle.setFont(font);
        newCellStyle.setWrapText(true);
        newCellStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        newCellStyle.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        newCellStyle.setLocked(true);
        newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //newCellStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        newCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        newCellStyle.setBorderBottom(BorderStyle.THIN);
        newCellStyle.setBorderLeft(BorderStyle.THIN);
        return newCellStyle;
    }

    public void buildCurrentSheet(com.alibaba.excel.metadata.Sheet sheet, Builder builder) {
        if (sheetMap.containsKey(sheet.getSheetNo())) {
            this.currentSheet = sheetMap.get(sheet.getSheetNo());
        } else {
            this.currentSheet = workbook.createSheet(
                sheet.getSheetName() != null ? sheet.getSheetName() : sheet.getSheetNo() + "");
            this.currentSheet.setDefaultColumnWidth(20);
            sheetMap.put(sheet.getSheetNo(), this.currentSheet);
            if (sheet.getHead() != null) {
                this.head = sheet.getHead();
            } else if (sheet.getClazz() != null) {
                this.head = buildHead(sheet.getClazz());
            }
            this.currentClazz = sheet.getClazz();
            buildTableStyle(sheet.getTableStyle());

            builder.createHead();

        }

    }

    private void buildTableStyle(TableStyle tableStyle) {
        if (tableStyle != null) {
            CellStyle headStyle = buildDefaultCellStyle();
            if (tableStyle.getTableHeadFont() != null) {
                Font font = this.workbook.createFont();
                font.setFontName(tableStyle.getTableHeadFont().getFontName());
                font.setFontHeightInPoints(tableStyle.getTableHeadFont().getFontHeightInPoints());
                font.setBold(tableStyle.getTableHeadFont().isBold());
                headStyle.setFont(font);
            }
            if (tableStyle.getTableHeadBackGroundColor() != null) {
                headStyle.setFillForegroundColor(tableStyle.getTableHeadBackGroundColor().getIndex());
            }
            this.currentHeadCellStyle = headStyle;
            CellStyle contentStyle = buildDefaultCellStyle();
            if (tableStyle.getTableContentFont() != null) {
                Font font = this.workbook.createFont();
                font.setFontName(tableStyle.getTableContentFont().getFontName());
                font.setFontHeightInPoints(tableStyle.getTableContentFont().getFontHeightInPoints());
                font.setBold(tableStyle.getTableContentFont().isBold());
                contentStyle.setFont(font);
            }
            if (tableStyle.getTableContentBackGroundColor() != null) {
                contentStyle.setFillForegroundColor(tableStyle.getTableContentBackGroundColor().getIndex());
            }
            this.currentContentCellStyle = contentStyle;
        }
    }

    public void buildTable(Table table, Builder builder) {
        if (!tableMap.containsKey(table.getTableNo())) {
            if (table.getHead() != null) {
                this.head = table.getHead();
            } else if (table.getClazz() != null) {
                this.head = buildHead(table.getClazz());
            }
            this.currentClazz = table.getClazz();
            tableMap.put(table.getTableNo(), table);
            buildTableStyle(table.getTableStyle());
            builder.createHead();
        }

    }

    private List<List<String>> buildHead(Class clazz) {
        Field[] fields = clazz.getDeclaredFields();
        List<List<String>> head = new ArrayList<List<String>>();
        for (int i = 0; i < fields.length; i++) {
            Field f = fields[i];
            ExcelProperty p = f.getAnnotation(ExcelProperty.class);
            if (p != null) {
                String[] value = p.value();
                head.add(Arrays.asList(value));
            }
        }
        return head;
    }

    public Class<?> getCurrentClazz() {
        return currentClazz;
    }

    public void setCurrentClazz(Class<?> currentClazz) {
        this.currentClazz = currentClazz;
    }

    public Sheet getCurrentSheet() {
        return currentSheet;
    }

    public void setCurrentSheet(Sheet currentSheet) {
        this.currentSheet = currentSheet;
    }

    public String getCurrentSheetName() {
        return currentSheetName;
    }

    public void setCurrentSheetName(String currentSheetName) {
        this.currentSheetName = currentSheetName;
    }

    public ExcelTypeEnum getExcelType() {
        return excelType;
    }

    public void setExcelType(ExcelTypeEnum excelType) {
        this.excelType = excelType;
    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public void setOutputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    public List<List<String>> getHead() {
        return head;
    }

    public void setHead(List<List<String>> head) {
        this.head = head;
    }

    public CellStyle getCurrentHeadCellStyle() {
        return this.currentHeadCellStyle == null ? defaultCellStyle : this.currentHeadCellStyle;
    }

    public CellStyle getCurrentContentStyle() {
        return this.currentContentCellStyle;
    }

    public List<Object> getCurrentData() {
        return currentData;
    }

    public void setCurrentData(List<Object> currentData) {
        this.currentData = currentData;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

}
