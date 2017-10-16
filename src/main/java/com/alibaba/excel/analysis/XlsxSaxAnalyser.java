package com.alibaba.excel.analysis;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.alibaba.excel.metadata.IndexValue;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.util.IndexValueConverter;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Created by jipengfei on 17/2/18.
 */
public class XlsxSaxAnalyser extends SaxAnalyser {

    private AnalysisContext analysisContext;

    private final int startRow;

    private int currentRow = 0;

    private List<IndexValue> rowData;

    @Override
    public List<Sheet> getSheets() {
        List<Sheet> sheets = new ArrayList<Sheet>();
        try {
            OPCPackage pkg = OPCPackage.open(analysisContext.getInputStream());
            XSSFReader xssfReader = new XSSFReader(pkg);
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator)xssfReader.getSheetsData();
            int i = 1;
            while (iter.hasNext()) {
                iter.next();
                Sheet sheet = new Sheet(i, 0);
                String sheetName = iter.getSheetName();
                sheet.setSheetName(sheetName);
                i++;
                sheets.add(sheet);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return sheets;
    }

    public void execute() {
        try {
            OPCPackage pkg = OPCPackage.open(analysisContext.getInputStream());
            XSSFReader r = new XSSFReader(pkg);

            SharedStringsTable sst = r.getSharedStringsTable();
            Sheet sheetParam = analysisContext.getCurrentSheet();
            InputStream sheetInputStream = null;
            if (sheetParam != null && sheetParam.getSheetNo() > 0) {
                sheetInputStream = r.getSheet("rId" + sheetParam.getSheetNo());
                InputSource sheetSource = new InputSource(sheetInputStream);
                getSheetParser(sst).parse(sheetSource);
                sheetInputStream.close();
            } else {
                Iterator<InputStream> ite = r.getSheetsData();
                while (ite.hasNext()) {
                    sheetInputStream = ite.next();
                    if (sheetInputStream != null) {
                        InputSource sheetSource = new InputSource(sheetInputStream);
                        getSheetParser(sst).parse(sheetSource);
                        sheetInputStream.close();
                    }
                }

            }

        } catch (Exception e) {
            throw new ExcelAnalysisException(e);
        }

    }

    public XlsxSaxAnalyser(AnalysisContext analysisContext) {

        this.analysisContext = analysisContext;
        this.startRow = 0;
        analysisContext.setCurrentRownNum(0);

    }

    private XMLReader getSheetParser(SharedStringsTable sst) throws SAXException, ParserConfigurationException {

        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        SAXParser saxParser = saxFactory.newSAXParser();
        XMLReader parser = saxParser.getXMLReader();
        ContentHandler handler = new PagingHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }

    private boolean isAccess() {
        if (currentRow >= startRow) {
            return true;
        }
        return false;
    }

    public void notifyListeners() {
        for (AnalysisEventListener listener : getAllRegister()) {
            listener.invoke(IndexValueConverter.converter(rowData), analysisContext);
        }
    }

    /**
     *
     */
    private class PagingHandler extends DefaultHandler {
        private SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;
        private String index = null;
        boolean isTElement = false;
        boolean notAllEmpty = false;

        private PagingHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        @Override
        public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
            if ("dimension".equals(name)) {
                //获得总计录数
                String d = attributes.getValue("ref");
                String totalStr =d.substring(d.indexOf(":")+1,d.length());

                String c = totalStr.toUpperCase().replaceAll("[A-Z]", "");
                analysisContext.setTotalCount(Integer.parseInt(c));
            }
            if (name.equals("c")) {
                index = attributes.getValue("r");
                if (index.contains("N")) {
                }
                if (isNewRow(index, currentRow)) {
                    if (rowData != null && isAccess() && !rowData.isEmpty() && notAllEmpty) {
                        analysisContext.setCurrentRownNum(currentRow-1);
                        notifyListeners();
                    }
                    rowData = new ArrayList<IndexValue>();
                    currentRow++;

                    notAllEmpty = false;
                }
                if (isAccess()) {
                    String cellType = attributes.getValue("t");
                    if (cellType != null && cellType.equals("s")) {
                        nextIsString = true;
                    } else {
                        nextIsString = false;
                    }
                }

            }
            if ("t".equals(name)) {
                isTElement = true;
            } else {
                isTElement = false;
            }
            lastContents = "";
        }

        private Boolean isNewRow(String idex, int currentRow) {
            String num = "";
            for (int i = 0; i < index.length(); i++) {
                char c = index.charAt(i);
                if (c <= '9' && c >= '0') {
                    num = num + c;
                }
            }
            int n = Integer.parseInt(num);
            if (n > currentRow) {
                return true;
            } else {
                return false;
            }
        }

        @Override
        public void endElement(String uri, String localName, String name) throws SAXException {
            if (isAccess()) {
                if (nextIsString) {
                    int idx = Integer.parseInt(lastContents);
                    lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                    nextIsString = false;
                }
                // 将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符
                if (lastContents != null) {
                    lastContents = lastContents.trim();
                }
                if (isTElement) {
                    if (lastContents != null && !"".equals(lastContents)) {
                        notAllEmpty = true;
                    }
                    rowData.add(new IndexValue(index, lastContents));
                }
                if (name.equals("v")) {
                    if (lastContents != null && !"".equals(lastContents)) {
                        notAllEmpty = true;
                    }
                    rowData.add(new IndexValue(index, lastContents));
                }
            }

        }

        @Override
        public void characters(char[] ch, int start, int length) throws SAXException {
            if (isAccess()) {
                lastContents += new String(ch, start, length);
            }

        }

        @Override
        public void endDocument() throws SAXException {
            if (rowData != null && isAccess() && !rowData.isEmpty()) {
                analysisContext.setCurrentRownNum(currentRow-1);
                notifyListeners();
            }

        }

    }

}
