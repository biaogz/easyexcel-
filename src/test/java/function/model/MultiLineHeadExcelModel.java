package function.model;

import com.alibaba.excel.annotation.ExcelColumnNum;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by jipengfei on 17/2/19.
 */
public class MultiLineHeadExcelModel extends BaseRowModel {

    @ExcelColumnNum(0)
    @ExcelProperty({"表头1","表头1","表头31"})
    private String p1;

    @ExcelColumnNum(1)
    @ExcelProperty({"表头1","表头1","表头32"})
    private String p2;

    @ExcelColumnNum(2)
    @ExcelProperty({"表头3","表头3","表头3"})
    private int p3;

    @ExcelColumnNum(3)
    @ExcelProperty({"表头4","表头4","表头4"})
    private long p4;

    @ExcelColumnNum(4)
    @ExcelProperty({"表头5","表头51","表头52"})
    private String p5;

    @ExcelColumnNum(5)
    @ExcelProperty({"表头6","表头61","表头611"})
    private String p6;

    @ExcelColumnNum(6)
    @ExcelProperty({"表头6","表头61","表头612"})
    private String p7;

    @ExcelColumnNum(7)
    @ExcelProperty({"表头6","表头62","表头621"})
    private String p8;

    @ExcelColumnNum(8)
    @ExcelProperty({"表头6","表头62","表头622"})
    private String p9;

    public String getP1() {
        return p1;
    }

    public void setP1(String p1) {
        this.p1 = p1;
    }

    public String getP2() {
        return p2;
    }

    public void setP2(String p2) {
        this.p2 = p2;
    }

    public int getP3() {
        return p3;
    }

    public void setP3(int p3) {
        this.p3 = p3;
    }

    public long getP4() {
        return p4;
    }

    public void setP4(long p4) {
        this.p4 = p4;
    }

    public String getP5() {
        return p5;
    }

    public void setP5(String p5) {
        this.p5 = p5;
    }

    public String getP6() {
        return p6;
    }

    public void setP6(String p6) {
        this.p6 = p6;
    }

    public String getP7() {
        return p7;
    }

    public void setP7(String p7) {
        this.p7 = p7;
    }

    public String getP8() {
        return p8;
    }

    public void setP8(String p8) {
        this.p8 = p8;
    }

    public String getP9() {
        return p9;
    }

    public void setP9(String p9) {
        this.p9 = p9;
    }

    public static void main(String[] args) {
        Class<?> clazz = MultiLineHeadExcelModel.class;
        Field[] fields = clazz.getDeclaredFields();
        List<List<String>> head = new ArrayList<List<String>>();
        for (int i = 0; i < fields.length ; i++) {
            Field f = fields[i];
            ExcelProperty p = f.getAnnotation(ExcelProperty.class);
            String[] value =p.value();
            head.add(Arrays.asList(value));
        }
        System.out.println(head);
    }
}
