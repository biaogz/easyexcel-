package function.read;

import java.io.InputStream;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.parameter.AnalysisParam;
import com.alibaba.excel.support.ExcelTypeEnum;

import function.listener.ExcelListener;
import function.model.TestModel3;
import org.junit.Test;

/**
 * Created by jipengfei on 17/3/19.
 *
 * @author jipengfei
 * @date 2017/03/19
 */
public class Test3 {

    @Test
    public void testExcel2007WithReflectModel() {
        InputStream inputStream = getInputStream("test3.xlsx");

        // 解析每行结果在listener中处理
        AnalysisEventListener listener = new ExcelListener();

        new ExcelReader(new AnalysisParam(inputStream, ExcelTypeEnum.XLSX, null),listener).read(new Sheet(1,1),TestModel3.class);

    }

    private InputStream getInputStream(String fileName) {
        return Thread.currentThread().getContextClassLoader().getResourceAsStream(""+fileName);

    }
}
