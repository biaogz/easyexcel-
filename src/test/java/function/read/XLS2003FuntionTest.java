package function.read;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.parameter.AnalysisParam;
import com.alibaba.excel.support.ExcelTypeEnum;

import function.listener.ExcelListener;
import function.model.LoanInfo;
import function.model.OneRowHeadExcelModel;
import junit.framework.TestCase;
import org.junit.Test;

import java.io.InputStream;

/**
 * Created by jipengfei on 17/2/19.
 */
public class XLS2003FuntionTest extends TestCase {

    @Test
    public void testExcel2003NoModel() {
        InputStream inputStream = getInputStream("loan1.xls");
        // 解析每行结果在listener中处理
        ExcelListener listener = new ExcelListener();

        new ExcelReader(new AnalysisParam(inputStream, ExcelTypeEnum.XLS, null), listener).read();

    }

    @Test
    public void testExcel2003WithSheet() {
        InputStream inputStream = getInputStream("loan1.xls");
        // 解析每行结果在listener中处理
        ExcelListener listener = new ExcelListener();
        new ExcelReader(new AnalysisParam(inputStream, ExcelTypeEnum.XLS, null), listener).read(new Sheet(1, 1));
        System.out.println(listener.getDatas());
    }

    @Test
    public void testExcel2003WithReflectModel() {
        InputStream inputStream = getInputStream("loan1.xls");

        // 解析每行结果在listener中处理
        AnalysisEventListener listener = new ExcelListener();

        new ExcelReader(new AnalysisParam(inputStream, ExcelTypeEnum.XLS, null),listener).read(new Sheet(1,2),LoanInfo.class);

    }

    private InputStream getInputStream(String fileName) {
        return Thread.currentThread().getContextClassLoader().getResourceAsStream("" + fileName);

    }
}
