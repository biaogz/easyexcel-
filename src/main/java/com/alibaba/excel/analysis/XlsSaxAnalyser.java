package com.alibaba.excel.analysis;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.alibaba.excel.metadata.Sheet;

import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NoteRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.RKRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.eventusermodel.XSSFReader;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by jipengfei on 17/2/18.
 */
public class XlsSaxAnalyser extends SaxAnalyser implements HSSFListener {

    private AnalysisContext analysisContext;

    public XlsSaxAnalyser(AnalysisContext context) {
        this.analysisContext = context;
        this.records = new ArrayList<String>();
        context.setCurrentRownNum(0);
        try {
            this.fs = new POIFSFileSystem(analysisContext.getInputStream());
        } catch (IOException e) {
            throw new ExcelAnalysisException(e);
        }
    }

    @Override
    public List<Sheet> getSheets() {
        execute();
        return sheets;
    }

    public void execute() {
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        formatListener = new FormatTrackingHSSFListener(listener);

        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();

        if (outputFormulaValues) {
            request.addListenerForAllRecords(formatListener);
        } else {
            workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener);
            request.addListenerForAllRecords(workbookBuildingListener);
        }

        try {
            factory.processWorkbookEvents(request, fs);
        } catch (IOException e) {
            throw new ExcelAnalysisException(e);
        }
    }

    public void notifyListeners() {
        for (AnalysisEventListener listener : getAllRegister()) {
            listener.invoke(records, analysisContext);
        }

    }

    private POIFSFileSystem fs;

    private int lastRowNumber;
    private int lastColumnNumber;

    /**
     * Should we output the formula, or the value it has?
     */
    private boolean outputFormulaValues = true;

    /**
     * For parsing Formulas
     */
    private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;
    private HSSFWorkbook stubWorkbook;

    // Records we pick up as we process
    private SSTRecord sstRecord;
    private FormatTrackingHSSFListener formatListener;

    /**
     * So we known which sheet we're on
     */
    // private BoundSheetRecord[] orderedBSRs;

    // @SuppressWarnings("rawtypes")
    // private ArrayList boundSheetRecords = new ArrayList();

    // For handling formulas with string results
    private int nextRow;
    private int nextColumn;
    private boolean outputNextStringRecord;

    /**
     * Main HSSFListener method, processes events, and outputs the CSV as the
     * file is processed.
     */

    private int sheetIndex;

    private List<String> records;

    boolean notAllEmpty = false;

    private BoundSheetRecord[] orderedBSRs;

    private List<BoundSheetRecord> boundSheetRecords = new ArrayList<BoundSheetRecord>();

    private List<Sheet> sheets = new ArrayList<Sheet>();

    public void processRecord(Record record) {
        int thisRow = -1;
        int thisColumn = -1;
        String thisStr = null;

        switch (record.getSid()) {
            case BoundSheetRecord.sid:
                boundSheetRecords.add((BoundSheetRecord)record);
                break;
            case BOFRecord.sid:
                BOFRecord br = (BOFRecord)record;
                if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
                    // Create sub workbook if required
                    if (workbookBuildingListener != null && stubWorkbook == null) {
                        stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
                    }

                    if (orderedBSRs == null) {
                        orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
                    }
                    sheetIndex++;
                    Sheet sheet = new Sheet(sheetIndex, 0);
                    sheet.setSheetName(orderedBSRs[sheetIndex - 1].getSheetname());
                    sheets.add(sheet);
                }
                break;

            case SSTRecord.sid:
                sstRecord = (SSTRecord)record;
                break;

            case BlankRecord.sid:
                BlankRecord brec = (BlankRecord)record;

                thisRow = brec.getRow();
                thisColumn = brec.getColumn();
                thisStr = "";
                break;
            case BoolErrRecord.sid:
                BoolErrRecord berec = (BoolErrRecord)record;

                thisRow = berec.getRow();
                thisColumn = berec.getColumn();
                thisStr = "";
                break;

            case FormulaRecord.sid:
                FormulaRecord frec = (FormulaRecord)record;

                thisRow = frec.getRow();
                thisColumn = frec.getColumn();

                if (outputFormulaValues) {
                    if (Double.isNaN(frec.getValue())) {
                        // Formula result is a string
                        // This is stored in the next record
                        outputNextStringRecord = true;
                        nextRow = frec.getRow();
                        nextColumn = frec.getColumn();
                    } else {
                        thisStr = formatListener.formatNumberDateCell(frec);
                    }
                } else {
                    thisStr = HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression());
                }
                break;
            case StringRecord.sid:
                if (outputNextStringRecord) {
                    // String for formula
                    StringRecord srec = (StringRecord)record;
                    thisStr = srec.getString();
                    thisRow = nextRow;
                    thisColumn = nextColumn;
                    outputNextStringRecord = false;
                }
                break;

            case LabelRecord.sid:
                LabelRecord lrec = (LabelRecord)record;

                thisRow = lrec.getRow();
                thisColumn = lrec.getColumn();
                thisStr = lrec.getValue();
                break;
            case LabelSSTRecord.sid:
                LabelSSTRecord lsrec = (LabelSSTRecord)record;

                thisRow = lsrec.getRow();
                thisColumn = lsrec.getColumn();
                if (sstRecord == null) {
                    thisStr = "";
                } else {
                    thisStr = sstRecord.getString(lsrec.getSSTIndex()).toString();
                }
                break;
            case NoteRecord.sid:
                NoteRecord nrec = (NoteRecord)record;

                thisRow = nrec.getRow();
                thisColumn = nrec.getColumn();
                // TODO: Find object to match nrec.getShapeId()
                thisStr = "(TODO)";
                break;
            case NumberRecord.sid:
                NumberRecord numrec = (NumberRecord)record;

                thisRow = numrec.getRow();
                thisColumn = numrec.getColumn();

                // Format
                thisStr = formatListener.formatNumberDateCell(numrec);
                break;
            case RKRecord.sid:
                RKRecord rkrec = (RKRecord)record;

                thisRow = rkrec.getRow();
                thisColumn = rkrec.getColumn();
                thisStr = "";
                break;
            default:
                break;
        }

        // Handle new row
        if (thisRow != -1 && thisRow != lastRowNumber) {
            lastColumnNumber = -1;
        }

        // Handle missing column
        if (record instanceof MissingCellDummyRecord) {
            MissingCellDummyRecord mc = (MissingCellDummyRecord)record;
            thisRow = mc.getRow();
            thisColumn = mc.getColumn();
            thisStr = "";
        }

        // If we got something to print out, do so
        if (thisStr != null) {
            //if (thisColumn > 1) {
            //    // output.print(',');
            //}
            //if (thisStr != null) {
            thisStr = thisStr.trim();
            if (!"".equals(thisStr)) {
                notAllEmpty = true;
            }
            //            }
            // output.print(thisStr);
            records.add(thisStr);
        }

        // Update column and row count
        if (thisRow > -1) {
            lastRowNumber = thisRow;
        }
        if (thisColumn > -1) {
            lastColumnNumber = thisColumn;
        }

        // Handle end of row
        if (record instanceof LastCellOfRowDummyRecord) {
            thisRow = ((LastCellOfRowDummyRecord)record).getRow();
            // thisColumn = ((LastCellOfRowDummyRecord)
            // record).getLastColumnNumber();
            // Columns are 0 based
            if (lastColumnNumber == -1) {
                lastColumnNumber = 0;
            }
            analysisContext.setCurrentRownNum(thisRow);
            Sheet sheet = analysisContext.getCurrentSheet();

            if ((sheet == null || sheet.getSheetNo() == sheetIndex) && notAllEmpty) {
                notifyListeners();
            }
            // System.out.println(records);
            records.clear();
            lastColumnNumber = -1;
            notAllEmpty = false;


        }
    }

}
