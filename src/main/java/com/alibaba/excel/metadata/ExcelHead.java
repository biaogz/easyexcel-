/*
s * Copyright 2016 Alibaba.com All right reserved. This software is the confidential and proprietary information of
 * Alibaba.com ("Confidential Information"). You shall not disclose such Confidential Information and shall use it only
 * in accordance with the terms of the license agreement you entered into with Alibaba.com.
 */
package com.alibaba.excel.metadata;

import java.util.ArrayList;
import java.util.List;

/**
 * ��ExcelHead.java��ʵ��������TODO ��ʵ������
 * 
 * @author jipengfei 2016��8��4�� ����12:35:47
 */
public class ExcelHead {

    //�ڲ�list����һ��
    private List<List<String>> head = new ArrayList<List<String>>();

    private int headRowNum = 0;
    

    public List<String> getLeaf() {
        List<String> leaf = new ArrayList<String>(head.size());
        for (List<String> list : head) {
            if (list!=null&&list.size()>0) {
                leaf.add(list.get(list.size() - 1));
            }
        }
        return leaf;
    }

    /**
     * ��ȡ��ͷ�м���
     * 
     * @return
     */
    public int getHeadRowNum() {

        return headRowNum;
    }

    /**
     * �����кţ���ѯ���еı�ͷ���ݰ����ϲ���Ԫ��
     * 
     * @param rowNum��0��ʼ
     * @return
     */
    public List<String> getHeadByRowNum(int rowNum) {
        List<String> l = new ArrayList<String>(head.size());
        for (List<String> list : head) {
            if (list.size() > rowNum) {
                l.add(list.get(rowNum));
            } else {
                l.add(list.get(list.size() - 1));
            }
        }
        return l;
    }

    /**
     * �����кŷ��ر�ͷ�㼶��
     * 
     * @param columnNum
     * @return
     */
    public List<String> getHeadTreeByColumnNum(int columnNum) {
        return head.get(columnNum);
    }

    public int size() {
        return head.size();
    }

    public ExcelHead(List<List<String>> head){
        this.head = head;
        setHeadRownNum();
    }

    private void setHeadRownNum() {
        for (List<String> list : head) {
            if (list!=null&&list.size()>0) {
                if (list.size() > headRowNum) {
                    headRowNum = list.size();
                }
            }
        }
    }

    public List<List<String>> getHead() {
        return head;
    }

    public void setHead(List<List<String>> head) {
        this.head = head;
        setHeadRownNum();
    }

    public class CellRangeModel {

       private int firstRow;
       private int lastRow;
       private int firstCol;
       private int lastCol;
        public CellRangeModel(int firstRow,int lastRow,int firstCol,int lastCol){
            this.firstRow=firstRow;
            this.lastRow=lastRow;
            this.firstCol=firstCol;
            this.lastCol=lastCol;
        }
        
        public int getFirstRow() {
            return firstRow;
        }
        
        public void setFirstRow(int firstRow) {
            this.firstRow = firstRow;
        }
        
        public int getLastRow() {
            return lastRow;
        }
        
        public void setLastRow(int lastRow) {
            this.lastRow = lastRow;
        }
        
        public int getFirstCol() {
            return firstCol;
        }
        
        public void setFirstCol(int firstCol) {
            this.firstCol = firstCol;
        }
        
        public int getLastCol() {
            return lastCol;
        }
        
        public void setLastCol(int lastCol) {
            this.lastCol = lastCol;
        }
        
    }

    public List<CellRangeModel> getCellRangeModels() {
        List<CellRangeModel> rangs= new ArrayList<CellRangeModel>();
        for (int i = 0; i < head.size(); i++) {//i������
            List<String> columnvalues = head.get(i);
            for (int j = 0; j < columnvalues.size(); j++) {//j������
                int lastRow = getLastRangRow(j,columnvalues.get(j),columnvalues);
                int lastColumn = geatLastRangColumn(columnvalues.get(j),getHeadByRowNum(j),i);
                if(lastRow>=0&&lastColumn>=0&&(lastRow>j||lastColumn>i)){
                    rangs.add(new CellRangeModel(j,lastRow,i,lastColumn));
                }
                
            }
        }
        return rangs;
    }

    /**
     *
     * @param value
     * @param headByRowNum
     * @param i
     * @return
     */
    private int geatLastRangColumn(String value, List<String> headByRowNum,int i) {
        if(headByRowNum.indexOf(value)<i){
            return -1;
        }else{
            return headByRowNum.lastIndexOf(value);
        }
    }


    private int getLastRangRow(int j,String value, List<String> columnvalue) {

        if(columnvalue.indexOf(value)<j){
            return -1;
        }
        if(value!=null&&value.equals(columnvalue.get(columnvalue.size()-1))){
            return headRowNum-1; 
        }else{
            return columnvalue.lastIndexOf(value); 
        }
    }

}
