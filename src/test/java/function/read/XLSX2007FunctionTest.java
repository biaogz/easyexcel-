package function.read;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.parameter.AnalysisParam;
import com.alibaba.excel.support.ExcelTypeEnum;

import function.listener.ExcelListener;
import function.model.OneRowHeadExcelModel;
import junit.framework.TestCase;
import org.junit.Test;

import java.io.InputStream;

/**
 * Created by jipengfei on 17/2/18.
 */
public class XLSX2007FunctionTest extends TestCase{

    //创建没有自定义模型,没有sheet的解析器,默认解析所有sheet解析结果以List<String>的方式通知监听者
    @Test
    public void testExcel2007NoModel() {
        InputStream inputStream = getInputStream("2007.xlsx");

        // 解析每行结果在listener中处理
        AnalysisEventListener listener = new ExcelListener();

        new ExcelReader(new AnalysisParam(inputStream, ExcelTypeEnum.XLSX, null),listener ).read();
    }

    //创建没有自定义模型,但有规定sheet解析器,解析结果以List<String>的方式通知监听者
    @Test
    public void testExcel2007WithSheet() {
        InputStream inputStream = getInputStream("test2.xlsx");

        // 解析每行结果在listener中处理
        AnalysisEventListener listener = new ExcelListener();

        new ExcelReader(new AnalysisParam(inputStream, ExcelTypeEnum.XLSX, null), listener).read(new Sheet(1,1));
    }

    //创建需要反射映射模型的解析器,解析结果List<Object> Object为自定义的模型
    @Test
    public void testExcel2007WithReflectModel() {
        InputStream inputStream = getInputStream("2007.xlsx");

        // 解析每行结果在listener中处理
        AnalysisEventListener listener = new ExcelListener();

        new ExcelReader(new AnalysisParam(inputStream, ExcelTypeEnum.XLSX, null),listener).read(new Sheet(1,1),OneRowHeadExcelModel.class);

    }

    @Test
    public void testExcel2007MultHeadWithReflectModel() {
        InputStream inputStream = getInputStream("2007_1.xlsx");

        // 解析每行结果在listener中处理
        AnalysisEventListener listener = new ExcelListener();

        new ExcelReader(new AnalysisParam(inputStream, ExcelTypeEnum.XLSX, null), listener).read(new Sheet(1,4),OneRowHeadExcelModel.class);


    }



    private InputStream getInputStream(String fileName) {
        return Thread.currentThread().getContextClassLoader().getResourceAsStream(""+fileName);

    }
}
