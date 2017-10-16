package function.read;

import java.io.InputStream;
import java.util.List;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.parameter.AnalysisParam;
import com.alibaba.excel.support.ExcelTypeEnum;

import function.listener.ExcelListener;
import function.model.TestModel3;
import org.junit.Test;

/**
 * Created by jipengfei on 17/3/22.
 *
 * @author jipengfei
 * @date 2017/03/22
 */
public class ReadSheets {
    @Test
    public void ReadSheets2007() {
        InputStream inputStream = getInputStream("2007.xlsx");

        // 解析每行结果在listener中处理
        AnalysisEventListener listener = new ExcelListener();

       ExcelReader reader=  new ExcelReader(new AnalysisParam(inputStream, ExcelTypeEnum.XLSX, null),listener);

       List<Sheet> sheets = reader.getSheets();
        System.out.println(sheets);

    }

    @Test
    public void ReadSheets2003() {
        InputStream inputStream = getInputStream("2003.xls");

        // 解析每行结果在listener中处理
        AnalysisEventListener listener = new ExcelListener();

        ExcelReader reader=  new ExcelReader(new AnalysisParam(inputStream, ExcelTypeEnum.XLS, null),listener);

        List<Sheet> sheets = reader.getSheets();
        System.out.println(sheets);

    }

    private InputStream getInputStream(String fileName) {
        return Thread.currentThread().getContextClassLoader().getResourceAsStream(""+fileName);

    }
}
