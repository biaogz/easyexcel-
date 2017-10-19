## easyexcel解决的问题
### 内存溢出
使用开源的框架来做解析，目前最稳定应该是apache的POI，但大部分使用POI都是使用他的userModel模式。userModel的好处是上手容易使用简单，随便拷贝个代码跑一下，剩下就是写业务转换了，虽然转换也要写上百行代码，但相对比较好理解。然而userModel模式最大的问题是它非常吃内存，一个几兆的文件解析要用掉上百兆的内存。现在很多应用采用这种模式，之所以还正常在跑一定是使用该功能人不多，多的话一定会爆掉的。
### SAX模式能够解决内存溢出，但使用复杂，部分数据格式不支持
对POI有过深入了解的估计才知道原来POI还有SAX模式。但SAX模式相对比较复杂，excel有03和07两种版本，两个版本数据存储方式截然不同，sax解析方式也各不一样。想要了解清楚这两种解析方式，才去写代码测试，估计两天时间是需要的。再加上即使解析完，要转换到自己业务模型还要很多繁琐的代码。总体下来感觉至少需要三天，由于代码复杂，后续维护成本巨大。

## xls、xlsx、csv格式分析
- xls是Microsoft Excel2007前excel的文件存储格式，实现原理是基于微软的ole db是微软com组件的一种实现，本质上也是一个微型数据库，由于微软的东西很多不开源，另外也已经被淘汰，了解它的细节意义不大，底层的编程都是基于微软的com组件去开发的。
- xlsx是Microsoft Excel2007后excel的文件存储格式，实现是基于openXml和zip技术。这种存储简单，安全传输方便，同时处理数据也变的简单。
- csv 我们可以理解为纯文本文件，可以被excel打开。他的格式非常简单，解析起来和解析文本文件一样。

## 核心原理
POI在处理Excel方面确实比较方便，但是当Excel数据量比较大的时候，使用POI处理就会导致OOM: Java heap space的错误，当有大量数据写入xlsx文件时，POI为我们提供了SXSSFWorkBook类来处理，这个类的处理机制是当内存中的数据条数达到一个极限数量的时候就flush这部分数据，再依次处理余下的数据，这个在大多数场景能够满足需求。当对一个存有大量数据的文件的xlsx文件进行读操作时，使用WorkBook处理就不行了，因为POI对文件是先将文件中的cell读入内存，生成一个树的结构（针对Excel中的每个sheet，使用TreeMap存储sheet中的行）。如果数据量比较大，则同样会产生java.lang.OutOfMemoryError: Java heap space错误。这个错误明显就是内存不足，虽然可以通过调整JVM堆区域内存大小来解决，但是一定的内存只能解决一定的数据量，当文件的大小并不确定的时候使用调整堆内存大小的方法并不能很优雅的解决这个问题，POI官方推荐使用“XSSF and SAX（event API）”方式来解决，即使用SAX对xlsx中的sheet.xml文件进行解析，然后根据MS的OOXML格式规范在sheet.xml,style.xml和sharedString.xml中解析数据。如果我们直接使用SAX对这些xml进行解析，则必须使用SAX中的三个方法startElement，characters和endElement解析到sheet.xml中的数据，再在style.xml中解析出格式，或者到sharedString.xml中解析数据，所以必须熟悉MS的OOXML规范。

分析清楚POI后要解决OOM有2个关键。
### 1、避免将全部全部数据一次加载到内存
采用sax模式一行一行解析，并将一行的解析结果以观察者的模式通知处理。
![基础模板1 (2).png](http://ata2-img.cn-hangzhou.img-pub.aliyun-inc.com/82bb195ac62532963b2364d2e4da23e5.png)
### 2、抛弃不重要的数据
Excel解析时候会包含样式，字体，宽度等数据，但这些数据是我们不关系的，如果将这部分数据抛弃可以大大降低内存使用。Excel中数据如下Style占了相当大的空间。
```
<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>12360</WindowHeight>
  <WindowWidth>25600</WindowWidth>
  <WindowTopX>0</WindowTopX>
  <WindowTopY>0</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
 </Styles>
 <Worksheet ss:Name="工作表1">
  <Table ss:ExpandedColumnCount="1" ss:ExpandedRowCount="1" x:FullColumns="1"
   x:FullRows="1" ss:DefaultColumnWidth="65" ss:DefaultRowHeight="15">
   <Row>
    <Cell><Data ss:Type="String">sdsdsd+A1</Data></Cell>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageLayoutZoom>0</PageLayoutZoom>
   <Selected/>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>
```
### 3、easyexcel核心功能
#### 1、读任意大小的03、07版Excel不会OOM
#### 2、读Excel自动通过注解，把结果映射为java模型
#### 3、读Excel支持多sheet
#### 4、读Excel时候是否对Excel内容做trim()增加容错
#### 5、写小量数据的03版Excel（不要超过2000行）
#### 6、写任意大07版Excel不会OOM
#### 7、写Excel通过注解将表头自动写入Excel
#### 8、写Excel可以自定义Excel样式 如：字体，加粗，表头颜色，数据内容颜色
#### 9、写Excel到多个不同sheet
#### 10、写Excel时一个sheet可以写多个Table
#### 11、写Excel时候自定义是否需要写表头
## 升级心得
之前自己写的excel解析工具放在ata上，陆续有同学咨询如何使用。当时自己写工具时候应用相对不多，测试用例，文档不是很详细。有同学咨询我就给解释下如何使用，但年前年后经常有同学咨询我说他们线上的excel解析报OOM了。因为使用同学的增多，觉得有必要再优化下了，最近抽了点时间将代码进行了重构，剃掉一些不常用的功能，让工具更加的轻量化，更加好用。

1、把excel解析时候同时做非空校验去掉了，因为觉得excel解析就是解析，校验应该是解析完用户自己选择如何校验。然而去掉后并不是不提供校验功能了，其实是有了更完美的解决方案。需要参数校验，可以参考下fastvalidator，这是我和AE一个同学搞的一个框架，性能是外部同类参数校验框架性能的10倍左右。如果需要时候ata上搜fastvalidator，有详细的使用介绍http://www.atatech.org/articles/68662

 2、重构也不再区分excel的解析方式，之前区分大excel和小excel两者采用不同解析方式。之前会有newLargeExcelReader，或者newLessExcelReader。这次统一改为了new ExcelReader就好，重构后统一采用large的方式解析。之所以做这样的改变一是使用同学经常问我两者的差别，二是受罗胖的跨年演讲的影响。（一个好的产品，有的时候并不是选择越多越好，你就直接告诉我那个最好就行了。比如手机苹果和安卓，我们发现安卓手机会有很多个性化的东西，可以定制桌面，主题。但iphone桌面就是那么单一，我告诉你这样就是最好的就OK了）作为excel解析工具，我把他当做一个产品，手机和电脑，再让其他同学使用的时候，感觉到的简单好用。既要提供优良的性能，又让使用者感到简单。之前的2中模式本质区别，一个是全部加载到内存去遍历。一个是以事件监听的者的方式通知接收者。



## 测试数据分析
![POI usermodel PK easyexcel(Excel 2003).png](http://ata2-img.cn-hangzhou.img-pub.aliyun-inc.com/02c4bfbbab99a649788523d04f84a42f.png)
![POI usermodel PK easyexcel(Excel 2007).png](http://ata2-img.cn-hangzhou.img-pub.aliyun-inc.com/f6a8a19ec959f0eb564e652de523fc9e.png)
![POI usermodel PK easyexcel(Excel 2003) (1).png](http://ata2-img.cn-hangzhou.img-pub.aliyun-inc.com/26888f7ea1cb8dc56db494926544edf7.png)
![POI usermodel PK easyexcel(Excel 2007) (1).png](http://ata2-img.cn-hangzhou.img-pub.aliyun-inc.com/4de1ac95bdfaa4b1870b224af4f4cb75.png)
从上面的性能测试可以看出easyexcel在解析耗时上比poiuserModel模式弱了一些。主要原因是我内部采用了反射做模型字段映射，中间我也加了cache，但感觉这点差距可以接受的。但在内存消耗上差别就比较明显了，easyexcel在后面文件再增大，内存消耗几乎不会增加了。但poi userModel就不一样了，简直就要爆掉了。想想一个excel解析200M，同时有20个人再用估计一台机器就挂了。

## 快速开始
### 二方包依赖
```
<dependency>
	<groupId>com.alibaba.shared</groupId>
	<artifactId>easyexcel</artifactId>
        <version>1.2.6</version>
</dependency>
```
### 读Excel
使用easyexcel解析03、07版本的Excel只是ExcelTypeEnum不同，其他使用完全相同，使用者无需知道底层解析的差异。
#### 无java模型直接把excel解析的每行结果以List&lt;String&gt;返回 在ExcelListener获取解析结果
读excel代码示例如下：
```
    @Test
    public void testExcel2003NoModel() {
        InputStream inputStream = getInputStream("loan1.xls");
        try {
            // 解析每行结果在listener中处理
            ExcelListener listener = new ExcelListener();

            ExcelReader excelReader = new ExcelReader(inputStream, ExcelTypeEnum.XLS, null, listener);
            excelReader.read();
        } catch (Exception e) {

        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
```
ExcelListener示例代码如下：
```
 /* 解析监听器，
 * 每解析一行会回调invoke()方法。
 * 整个excel解析结束会执行doAfterAllAnalysed()方法
 *
 * 下面只是我写的一个样例而已，可以根据自己的逻辑修改该类。
 * @author jipengfei
 * @date 2017/03/14
 */
public class ExcelListener extends AnalysisEventListener {

    //自定义用于暂时存储data。
    //可以通过实例获取该值
    private List<Object> datas = new ArrayList<Object>();
    public void invoke(Object object, AnalysisContext context) {
        System.out.println("当前行："+context.getCurrentRowNum());
        System.out.println(object);
        datas.add(object);//数据存储到list，供批量处理，或后续自己业务逻辑处理。
        doSomething(object);//根据自己业务做处理
    }
    private void doSomething(Object object) {
        //1、入库调用接口
    }
    public void doAfterAllAnalysed(AnalysisContext context) {
       // datas.clear();//解析结束销毁不用的资源
    }
    public List<Object> getDatas() {
        return datas;
    }
    public void setDatas(List<Object> datas) {
        this.datas = datas;
    }
}
```
#### 有java模型映射 
java模型写法如下：
```
public class LoanInfo extends BaseRowModel {
    @ExcelProperty(index = 0)
    private String bankLoanId;
    
    @ExcelProperty(index = 1)
    private Long customerId;
    
    @ExcelProperty(index = 2,format = "yyyy/MM/dd")
    private Date loanDate;
    
    @ExcelProperty(index = 3)
    private BigDecimal quota;
    
    @ExcelProperty(index = 4)
    private String bankInterestRate;
    
    @ExcelProperty(index = 5)
    private Integer loanTerm;
    
    @ExcelProperty(index = 6,format = "yyyy/MM/dd")
    private Date loanEndDate;
    
    @ExcelProperty(index = 7)
    private BigDecimal interestPerMonth;

    @ExcelProperty(value = {"一级表头","二级表头"})
    private BigDecimal sax;
}
```
@ExcelProperty(index = 3)数字代表该字段与excel对应列号做映射，也可以采用 @ExcelProperty(value = {"一级表头","二级表头"})用于解决不确切知道excel第几列和该字段映射，位置不固定，但表头的内容知道的情况。
```
    @Test
    public void testExcel2003WithReflectModel() {
        InputStream inputStream = getInputStream("loan1.xls");
        try {
            // 解析每行结果在listener中处理
            AnalysisEventListener listener = new ExcelListener();

            ExcelReader excelReader = new ExcelReader(inputStream, ExcelTypeEnum.XLS, null, listener);

            excelReader.read(new Sheet(1, 2, LoanInfo.class));
        } catch (Exception e) {

        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }
```
带模型解析与不带模型解析主要在构造new Sheet(1, 2, LoanInfo.class)时候包含class。Class需要继承BaseRowModel暂时BaseRowModel没有任何内容，后面升级可能会增加一些默认的数据。
### 写Excel
#### 每行数据是List&lt;String&gt;无表头
```
  OutputStream out = new FileOutputStream("/Users/jipengfei/77.xlsx");
        try {
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX,false);
            //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
            Sheet sheet1 = new Sheet(1, 0);
            sheet1.setSheetName("第一个sheet");
            writer.write(getListString(), sheet1);
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
```
#### 每行数据是一个java模型有表头----表头层级为一
生成Excel格式如下图
![屏幕快照 2017-06-02 上午9.49.39.png](http://ata2-img.cn-hangzhou.img-pub.aliyun-inc.com/dfcb44d05380e2e26bce93f850d9fc99.png)

模型写法如下：
```
public class ExcelPropertyIndexModel extends BaseRowModel {

    @ExcelProperty(value = "姓名" ,index = 0)
    private String name;

    @ExcelProperty(value = "年龄",index = 1)
    private String age;

    @ExcelProperty(value = "邮箱",index = 2)
    private String email;

    @ExcelProperty(value = "地址",index = 3)
    private String address;

    @ExcelProperty(value = "性别",index = 4)
    private String sax;

    @ExcelProperty(value = "高度",index = 5)
    private String heigh;

    @ExcelProperty(value = "备注",index = 6)
    private String last;
}
```
   @ExcelProperty(value = "姓名",index = 0) value是表头数据，默认会写在excel的表头位置，index代表第几列。
```
 @Test
    public void test1() throws FileNotFoundException {
        OutputStream out = new FileOutputStream("/Users/jipengfei/78.xlsx");
        try {
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
            //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
            Sheet sheet1 = new Sheet(1, 0,ExcelPropertyIndexModel.class);
            writer.write(getData(), sheet1);
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
```
#### 每行数据是一个java模型有表头----表头层级为多层级
生成Excel格式如下图：
![屏幕快照 2017-06-02 上午9.53.07.png](http://ata2-img.cn-hangzhou.img-pub.aliyun-inc.com/0cdb1673665e7940cd670871afb4b3d7.png)
java模型写法如下：
```
public class MultiLineHeadExcelModel extends BaseRowModel {

    @ExcelProperty(value = {"表头1","表头1","表头31"},index = 0)
    private String p1;

    @ExcelProperty(value = {"表头1","表头1","表头32"},index = 1)
    private String p2;

    @ExcelProperty(value = {"表头3","表头3","表头3"},index = 2)
    private int p3;

    @ExcelProperty(value = {"表头4","表头4","表头4"},index = 3)
    private long p4;

    @ExcelProperty(value = {"表头5","表头51","表头52"},index = 4)
    private String p5;

    @ExcelProperty(value = {"表头6","表头61","表头611"},index = 5)
    private String p6;

    @ExcelProperty(value = {"表头6","表头61","表头612"},index = 6)
    private String p7;

    @ExcelProperty(value = {"表头6","表头62","表头621"},index = 7)
    private String p8;

    @ExcelProperty(value = {"表头6","表头62","表头622"},index = 8)
    private String p9;
}
```
写Excel写法同上，只需将ExcelPropertyIndexModel.class改为MultiLineHeadExcelModel.class


#### 一个Excel多个sheet写法
```
 @Test
    public void test1() throws FileNotFoundException {

        OutputStream out = new FileOutputStream("/Users/jipengfei/77.xlsx");
        try {
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX,false);
            //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
            Sheet sheet1 = new Sheet(1, 0);
            sheet1.setSheetName("第一个sheet");
            writer.write(getListString(), sheet1);

            //写第二个sheet sheet2  模型上打有表头的注解，合并单元格
            Sheet sheet2 = new Sheet(2, 3, MultiLineHeadExcelModel.class, "第二个sheet", null);
            sheet2.setTableStyle(getTableStyle1());
            writer.write(getModeldatas(), sheet2);

            //写sheet3  模型上没有注解，表头数据动态传入
            List<List<String>> head = new ArrayList<List<String>>();
            List<String> headCoulumn1 = new ArrayList<String>();
            List<String> headCoulumn2 = new ArrayList<String>();
            List<String> headCoulumn3 = new ArrayList<String>();
            headCoulumn1.add("第一列");
            headCoulumn2.add("第二列");
            headCoulumn3.add("第三列");
            head.add(headCoulumn1);
            head.add(headCoulumn2);
            head.add(headCoulumn3);
            Sheet sheet3 = new Sheet(3, 1, NoAnnModel.class, "第三个sheet", head);
            writer.write(getNoAnnModels(), sheet3);
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
```
#### 一个sheet中有多个表格
```
@Test
    public void test2() throws FileNotFoundException {
        OutputStream out = new FileOutputStream("/Users/jipengfei/77.xlsx");
        try {
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX,false);

            //写sheet1  数据全是List<String> 无模型映射关系
            Sheet sheet1 = new Sheet(1, 0);
            sheet1.setSheetName("第一个sheet");
            Table table1 = new Table(1);
            writer.write(getListString(), sheet1, table1);
            writer.write(getListString(), sheet1, table1);

            //写sheet2  模型上打有表头的注解
            Table table2 = new Table(2);
            table2.setTableStyle(getTableStyle1());
            table2.setClazz(MultiLineHeadExcelModel.class);
            writer.write(getModeldatas(), sheet1, table2);

            //写sheet3  模型上没有注解，表头数据动态传入,此情况下模型field顺序与excel现实顺序一致
            List<List<String>> head = new ArrayList<List<String>>();
            List<String> headCoulumn1 = new ArrayList<String>();
            List<String> headCoulumn2 = new ArrayList<String>();
            List<String> headCoulumn3 = new ArrayList<String>();
            headCoulumn1.add("第一列");
            headCoulumn2.add("第二列");
            headCoulumn3.add("第三列");
            head.add(headCoulumn1);
            head.add(headCoulumn2);
            head.add(headCoulumn3);
            Table table3 = new Table(3);
            table3.setHead(head);
            table3.setClazz(NoAnnModel.class);
            table3.setTableStyle(getTableStyle2());
            writer.write(getNoAnnModels(), sheet1, table3);
            writer.write(getNoAnnModels(), sheet1, table3);

            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
```
## BUG fix记录
1.0.1----完善测试用例，防止歧义，模型字段映射不上时候有抛异常，改为提醒。
1.0.2-----修复拿到一行数据后，存到list中，但最后处理时候变为空的bug。
1.0.3-----修复无@ExcelProperty标注的多余字段时候报错。
1.0.4-----修复日期类型转换时候数字问题。基础模型支持字段类型int,long,double,boolean,date,string
1.0.5----优化类型转换的性能。
1.0.6----增加@ExcelColumnNum,修复字符串前后空白，增加过滤功能。
1.0.8-----如果整行excel数据全部为空，则不解析返回。完善多sheet的解析。
1.0.9-----修复excel超过16列被覆盖的问题，修复数据只有一行时候无法透传的bug。

## 参考文章
apache poi 官方:http://poi.apache.org/spreadsheet/how-to.html#xssf_sax_api
XLS2CSV:http://www.docjar.com/html/api/org/apache/poi/hssf/eventusermodel/examples/XLS2CSVmra.java.html
XLSX2CSV:https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xssf/eventusermodel/XLSX2CSV.java

## 异步处理以及异常的反馈
由于excel解析和进度可以独立开发，进度作为单独的工具使用不仅excel解析可以使用，文件上传，批量数据处理，任务调用，都可以列为进度任务，进度如何处理欢迎使用自己另外的工具，进度工具地址：[进度工具](http://www.atatech.org/articles/61887)
欢迎加入一起交流
![IMG_3224.jpg](http://ata2-img.cn-hangzhou.img-pub.aliyun-inc.com/39581c8ac3e197cbcd3354816f95b892.jpg)
