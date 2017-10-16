package function.write;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.parameter.GenerateParam;

import function.model.MultiLineHeadExcelModel;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by jipengfei on 17/2/19.
 */
@Deprecated
public class GenerateExcelTest {

    //写excel
    @Test
    public void testCreate(){

        ExcelWriter writer = null;
        try {
            writer =
                new ExcelWriter(new GenerateParam("66", MultiLineHeadExcelModel.class,new FileOutputStream("/Users/jipengfei/77.xlsx")));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        writer.write(getdatas());
        writer.write(getdatas());
        List<String> list = new ArrayList<String>();
        list.add("oooo");list.add("oooo");
        List<List> ll = new ArrayList<List>();
        ll.add(list);
        writer.write(ll);
        writer.finish();
    }


    private List<MultiLineHeadExcelModel> getdatas(){
        List<MultiLineHeadExcelModel> MODELS = new ArrayList<MultiLineHeadExcelModel>();
        MultiLineHeadExcelModel model1 = new MultiLineHeadExcelModel();
        model1.setP1("111");
        model1.setP2("111");
        model1.setP3(11);model1.setP4(9);model1.setP5("111");model1.setP6("111");model1.setP7("111");model1.setP8("111");


        MultiLineHeadExcelModel model2 = new MultiLineHeadExcelModel();
        model2.setP1("111");
        model2.setP2("111");
        model2.setP3(11);model2.setP4(9);model2.setP5("111");model2.setP6("111");model2.setP7("111");model2.setP8("111");



        MultiLineHeadExcelModel model3 = new MultiLineHeadExcelModel();
        model3.setP1("111");
        model3.setP2("111");
        model3.setP3(11);model3.setP4(9);model3.setP5("111");model3.setP6("111");model3.setP7("111");model3.setP8("111");


        MODELS.add(model1); MODELS.add(model2); MODELS.add(model3);

        return MODELS;
    }
}
