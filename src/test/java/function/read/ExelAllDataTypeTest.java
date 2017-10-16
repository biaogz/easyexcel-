package function.read;

import java.io.InputStream;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.parameter.AnalysisParam;
import com.alibaba.excel.support.ExcelTypeEnum;

import function.listener.ExcelListener;
import function.model.AllDataTypeModel;
import junit.framework.TestCase;
import org.junit.Test;

/**
 * Created by jipengfei on 17/3/15.
 *
 * @author jipengfei
 * @date 2017/03/15
 */
public class ExelAllDataTypeTest extends TestCase {
    // 创建没有自定义模型,没有sheet的解析器,默认解析所有sheet解析结果以List<String>的方式通知监听者
    @Test
    public void testExcel2007WithReflectModel() {
        InputStream inputStream = getInputStream("test2.xlsx");

        // 解析每行结果在listener中处理
        AnalysisEventListener listener = new ExcelListener();

        new ExcelReader(new AnalysisParam(inputStream, ExcelTypeEnum.XLSX, null),listener).read(new Sheet(1,1),AllDataTypeModel.class);

    }

    private InputStream getInputStream(String fileName) {
        return Thread.currentThread().getContextClassLoader().getResourceAsStream("" + fileName);

    }
}
