# maven使用方式
```xml
<!-- excel导入导出 -->
<dependency>
    <groupId>com.github.stupdit1t</groupId>
    <artifactId>poi-excel</artifactId>
    <version>2.0.2</version>
</dependency>
```
## 一. 项目优势
- 简单快速上手，且满足绝大多数业务场景
- 屏蔽POI细节，学习成本低。
- 未使用注解方式实现，纯编码代码块，去除烦人的各种POJO
- 功能强大，导入支持严格的单元格校验，导出支持公式，复杂表头和尾部设计，以及单元格样式自定义支持
- 支持读取复杂Excel模板,替换变量输出Excel,变量用${}替代
## 二. 更新记录
>  有需求才有进步，这个轮子本身就是从0开始因为需求慢慢叠加起来的。有新需求提出来,我觉得合适会更新. 如有疑问可加群帮解答: 811606008

### v2.0.2 (不兼容1.x)
1. 修复数字转列字符, 列字符转数字不正确的BUG
2. 导入代码优化结构调整,其他代码优化


### v1.8.2
1. 修复合并单元格，边框线有时候不填充的BUG

### v1.8.1
   1. 导出修复BigDecimal和Float识别为字符串，不能应用公式的BUG
   2. 修复图片导出不同版本报错的BUG

### v1.8
   1. 添加单元格设置批注功能，方法在Column.comment  也支持回调设置，同样方法comment
   2. 修改导出xls还是xlsx设置选项到ExportRule中，默认xlsx

### v1.7
   1. 导入抽象规则类修改
   2. 添加读取Excel的方法readSheet,方便将Excel直接读取为Map
   
### v1.6
   1. 新增单元格支持函数导出, 使用方式如设置字段内容为 =SUM(A1:12), 具体函数参考Excel写法
   2. 新增读取Excel模板, 替换模板里的变量并输出Excel, 变量标记为${}
   3. POI版本升级至4.1.2

### v1.5

1. 新增单元格整体样式自定义功能, 可设置全局表头/标题/单元格样式

## 三. 导出

##### 选择xls还是xlsx？

> xls速度较快，单sheet最大65535行，体积大. xlsx速度慢，单sheet最大1048576行，体积小

##### 1. 简单导出

* 代码示例

```java
/**
 * 简单导出
 *
 * @throws Exception
 */
public static void simpleExport(){
        // 1.导出的header标题设置
    String[]headers={"项目名称","项目图","所属区域","省份","市","项目所属人","项目领导人","得分","平均分","创建时间"};
    // 2.导出header对应的字段设置
    Column[]columns={
        Column.field("projectName"),
        Column.field("img"),
        Column.field("areaName"),
        Column.field("province"),
        Column.field("city").width(3),
        Column.field("people"),
        Column.field("leader"),
        Column.field("scount"),
        Column.field("avg"),
        Column.field("createTime").datePattern("yyyy-MM-dd")
    };
    // 3.执行导出
    ExcelUtils.export(outPath,data,ExportRules.simpleRule(columns,headers));
}
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/146552540-fc5d311a-92c2-4abb-9814-699251e16b17.png)

##### 2. 简单导出 + 添加自定义属性

* 代码示例

```java
/**
 * 简单导出2
 *
 * @throws Exception
 */
public static void simpleExport2(){
    // 1.导出的header标题设置
    String[]headers={"项目名称","项目图","所属区域","省份","市","项目所属人","项目领导人","得分","平均分","创建时间"};
    // 2.导出header对应的字段设置
    Column[]columns={
        // 不设置宽度自适应
        Column.field("projectName"),
        // 4.9项目图片
        Column.field("img"),
        // 4.1设置此列宽度为10, 添加注释
        Column.field("areaName").width(10).comment("你好吗"),
        // 4.2设置此列下拉框数据
        Column.field("province").dorpDown(new String[]{"陕西省","山西省","辽宁省"}),
        // 4.3设置此列水平居右
        Column.field("city").align(HorizontalAlignment.RIGHT),
        // 4.4 设置此列垂直居上
        Column.field("people").valign(VerticalAlignment.TOP),
        // 4.5 设置此列单元格 自定义校验 只能输入文本
        Column.field("leader")
        .width(4),
        //.verifyCustom("LEN(G4)>2", "名字长度必须大于2位"),
        // 4.6设置此列单元格 整数 数据校验 ，同时设置背景色为棕色
        Column.field("scount")
        .verifyIntNum("10~2000")
        .backColor(IndexedColors.BROWN),
        // 4.7设置此列单元格 浮点数 数据校验， 同时设置字体颜色红色
        Column.field("avg").
        verifyFloatNum("0.0~20.0")
        .color(IndexedColors.RED),
        // 4.8设置此列单元格 日期 数据校验 ，同时宽度为20、限制用户表格输入、水平居中、垂直居中、背景色、字体颜色
        Column.field("createTime")
        .datePattern("yyyy-MM-dd")
        .verifyDate("2000-01-01~2020-12-12")
        .align(HorizontalAlignment.LEFT)
        .valign(VerticalAlignment.CENTER)
        .backColor(IndexedColors.YELLOW)
        .color(IndexedColors.GOLD),
    };
    // 3.尾部合计行设计
    Map<String, String> footerRules=new HashMap<>();
    footerRules.put("1,1,A,H","合计");
    footerRules.put("1,1,I,I",String.format("=SUM(I3:I%s)",2+data.size()));
    footerRules.put("1,1,J,J",String.format("=AVERAGE(J3:I%s)",2+data.size()));
    footerRules.put("1,1,K,K","作者:625");

    // 4.自定义header样式
    ICellStyle headerStyle=new ICellStyle(){
        @Override
        public CellPosition getPosition(){
                return CellPosition.HEADER;
        }

        @Override
        public void handleStyle(Font font,CellStyle cellStyle){
            // 加粗
            font.setBold(true);
            // 黑体
            font.setFontName("黑体");
            // 字号12
            font.setFontHeightInPoints((short)12);
            // 字体红色
            font.setColor(IndexedColors.RED.getIndex());
            // 背绿色
            cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            // 边框
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            // 居中
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 折行
            cellStyle.setWrapText(true);
        }
    };
    ExportRules exportRules=ExportRules.simpleRule(columns,headers)
    // 大标题
    .title("简单导出")
    // 自动序号
    .autoNum(true)
    // sheet名称
    .sheetName("简单导出")
    // 尾部合计行设计
    .footerRules(footerRules)
    // 导出格式定义
    .xlsx(true)
    // 自定义全局样式
    .globalStyle(headerStyle);
    // 5.执行导出
    ExcelUtils.export(outPath,data,exportRules,(fieldName,value,row,col)->{
            System.out.print("[打印] 字段:"+fieldName);
            System.out.print(" 字段值:"+value);
            System.out.print(" 行数据:"+row);
            System.out.println(" 单元格样式:"+col);
            // 设置当前单元格值
            return value;
        }
    );
}
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/146552597-44997d69-d408-42c6-9950-f3395b324aa8.png)


##### 3. 复杂表头导出

* 代码示例

```java
/**
 * 复杂导出
 *
 * @throws Exception
 */
public static void complexExport(){
    // 1.表头设置,可以对应excel设计表头，一看就懂
    HashMap<String, String> headerRules=new HashMap<>();
    headerRules.put("1,1,A,K","项目资源统计");
    headerRules.put("2,3,A,A","序号");
    headerRules.put("2,2,B,E","基本信息");
    headerRules.put("3,3,B,B","项目名称");
    headerRules.put("3,3,C,C","所属区域");
    headerRules.put("3,3,D,D","省份");
    headerRules.put("3,3,E,E","市");
    headerRules.put("2,3,F,F","项目所属人");
    headerRules.put("2,3,G,G","市项目领导人");
    headerRules.put("2,2,H,I","分值");
    headerRules.put("3,3,H,H","得分");
    headerRules.put("3,3,I,I","平均分");
    headerRules.put("2,3,J,J","创建时间");
    headerRules.put("2,3,K,K","项目图片");
    // 2.尾部设置，一般可以用来设计合计栏
    HashMap<String, String> footerRules=new HashMap<>();
    footerRules.put("1,2,A,C","合计:");
    footerRules.put("1,2,D,K","=SUM(H4:H13)");
    // 3.导出header对应的字段设置
    Column[]column={
    Column.field("projectName"),
    Column.field("areaName"),
    Column.field("province"),
    Column.field("city"),
    Column.field("people"),
    Column.field("leader"),
    Column.field("scount"),
    Column.field("avg"),
    Column.field("createTime"),
    Column.field("img")

    };
    // 4.执行导出到工作簿
    ExportRules exportRules=ExportRules.complexRule(column,headerRules)
    .footerRules(footerRules)
    .autoNum(true);
    ExcelUtils.export(outPath,data,exportRules);
}
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/146559368-092741f8-a00a-4ffb-bb3b-87a6a38f90b4.png)



##### 4. 复杂的对象级联导出

* 代码示例

```java
/**
 * 复杂的对象级联导出
 *
 * @throws Exception
 */
public static void complexExport2(){
        // 1.导出的header设置
    String[]header={"學生姓名","所在班級","所在學校","更多父母姓名"};
        // 2.导出header对应的字段设置，列宽设置
    Column[]column={
        Column.field("name"),
        Column.field("classRoom.name"),
        Column.field("classRoom.school.name"),
        Column.field("moreInfo.parent.age"),
    };
    // 3.执行导出到工作簿
    ExcelUtils.export(outPath,complexData,ExportRules.simpleRule(column,header));
}
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/146556910-420b4993-99eb-474f-b880-49bdefae5f97.png)


##### 4. map数据导出

```java
/**
 * map数据导出
 *
 * @throws Exception
 */
public static void mapExport(){
    // 1.导出的header设置
    String[]header={"姓名","年龄"};
    // 2.导出header对应的字段设置，列宽设置
        Column[]column={
        Column.field("name"),
        Column.field("age"),
    };
    // 3.执行导出到工作簿
    ExcelUtils.export(outPath,mapData,ExportRules.simpleRule(column,header));
}
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/146559404-cc6c2bac-4236-453c-ba10-6b7cc97182ca.png)


##### 5. 模板导出

* 代码示例

```java
/**
 * 模板导出
 *
 * @throws Exception
 */
public static void templateExport(){
    // 1.导出的header设置
    String[]header={"宝宝姓名","宝宝昵称","家长姓名","手机号码","宝宝生日","月龄","宝宝性别","来源渠道","市场人员","咨询顾问","客服顾问","分配校区","备注"};
    // 2.导出header对应的字段设置，列宽设置
    Column[]column={Column.field("宝宝姓名"),Column.field("宝宝昵称"),
    Column.field("家长姓名"),
    Column.field("手机号码").verifyText("11~11","请输入11位的手机号码！"),
    Column.field("宝宝生日").datePattern("yyyy-MM-dd").verifyDate("2000-01-01~3000-12-31"),
    Column.field("月龄").width(4).verifyCustom("VALUE(F3:F6000)","月齡格式：如1年2个月则输入14"),
    Column.field("宝宝性别").dorpDown(new String[]{"男","女"}),
    Column.field("来源渠道").width(12).dorpDown(new String[]{"品推","市场"}),
    Column.field("市场人员").width(6).dorpDown(new String[]{"张三","李四"}),
    Column.field("咨询顾问").width(6).dorpDown(new String[]{"张三","李四"}),
    Column.field("客服顾问").width(6).dorpDown(new String[]{"大唐","银泰"}),
    Column.field("分配校区").width(6).dorpDown(new String[]{"大唐","银泰"}),
    Column.field("备注")
    };
    // 3.执行导出到工作簿
    ExcelUtils.export(outPath,Collections.emptyList(),ExportRules.simpleRule(column,header));
}
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/146559476-d16b4dd5-bb91-4971-ac06-cbe6e548599d.png)


##### 6. 支持多sheet导出

* 代码示例

```java
// 1.创建空workbook
Workbook emptyWorkbook=ExcelUtils.createEmptyWorkbook(true);
// 2.填充3个sheet数据
ExcelUtils.fillBook(emptyWorkbook,data1,ExportRules.simpleRule(column1,header1).sheetName("sheet1"));
ExcelUtils.fillBook(emptyWorkbook,data2,ExportRules.simpleRule(column2,header2).sheetName("sheet2"));
ExcelUtils.fillBook(emptyWorkbook,data3,ExportRules.simpleRule(column3,header3).sheetName("sheet3"));
// 3.导出
emptyWorkbook.write(new FileOutputStream(outPath));

```
* 导出示例

![image](https://user-images.githubusercontent.com/29246805/146565288-e711bff8-5f2f-4cda-98fc-3dfad2ba9baa.png)


##### 7. 支持大数据内存导出，防止OOM

* 代码示例

```java
// 1.声明大数据内存导出
Workbook emptyWorkbook=ExcelUtils.createBigWorkbook();
// 2.填充数据
ExcelUtils.fillBook(emptyWorkbook,data,ExportRules.simpleRule(column,header));
// 3.导出
emptyWorkbook.write(new FileOutputStream(outPath));
```

##### 8. 读模板替换变量导出
![image](https://user-images.githubusercontent.com/29246805/146565796-5fec5955-6356-49d6-b505-e6ed71e61127.png)
        
* 代码示例

```java
/**
 * 读模板替换变量导出
 *
 * @throws Exception
 */
public static void readExport(){
    Map<String, String> params=new HashMap<>();
    params.put("author","625");
    params.put("text","合计");
    params.put("area","西安市");
    params.put("prov","陕西省");
    Workbook workbook=ExcelUtils.readExcelWrite(templatePath,params);
    try{
        workbook.write(new FileOutputStream(outPath));
    }catch(IOException e){
        e.printStackTrace();
    }
}
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/146565886-0eea48f3-481a-4cc5-a3a2-afa230c50ff6.png)


## 五. 导入
##### 1. 支持严格的单元格校验,可以定位到单元格坐标校验
##### 2. 支持数据行的图片导入
##### 3. 支持导入过程中,对数据处理添加回调逻辑,满足其他业务场景
##### 4. xls和xlsx都支持导入

* 导入文件示例图

![image](https://user-images.githubusercontent.com/29246805/146621557-0732b698-dc2a-49c7-9ce4-fd4a6ce7bb36.png)


* 数据要转换的Java对象

```java
public class DemoData {

    private BigDecimal bigDecimalHandler;

    private Boolean booleanHandler;

    private Character charHandler;

    private Date dateHandler;

    private Double doubleHandler;

    private Float floatHandler;

    private Integer integerHandler;

    private Long longHandler;

    private Map<String, Object> objectHandler;

    private byte[] imgHandler;

    private String stringHandler;

    private Short shortHandler;
}
```

* 导入方法

```java
public class MainClass {

    public static void main(String[] args) {
        parseSheet();
    }

    public static void parseSheet() {
        // 1. 导入规则定义
        Consumer<AbsSheetVerifyRule> columnRule = (rule) -> {
            // 表示C列数据提取到字段bigDecimalHandler,字段为BigDecimal类型, 并且列不能为空
            rule.addRule("C", "bigDecimalHandler", "名字", new BigDecimalHandler(false, value -> {
                // 自定义处理, 名字不能是1.2345
                if (new BigDecimal(String.valueOf(value)).equals(new BigDecimal("1.2345"))) {
                    throw PoiException.error("不能是1.2345");
                }
                return new BigDecimal(String.valueOf(value));
            }));
            rule.addRule("E", "booleanHandler", "布尔宝", new BooleanHandler(true));
            rule.addRule("G", "charHandler", "char宝", new CharHandler(true));
            // 日期处理格式化,日期可以是 数字 或 字符串 或 excel日期
            rule.addRule("I", "dateHandler", "日期宝", new DateHandler("yyyy-MM-dd HH:mm:ss", true));
            rule.addRule("K", "doubleHandler", "double宝", new DoubleHandler(true));
            rule.addRule("M", "floatHandler", "float宝", new FloatHandler(true));
            rule.addRule("O", "integerHandler", "integer宝", new IntegerHandler(true));
            rule.addRule("G", "longHandler", "long宝", new LongHandler(true));
            // 数值转换对象或者枚举字典等等处理
            rule.addRule("S", "objectHandler", "对象宝", new ObjectHandler(true, (value) -> {
                Map<String, Object> map = new HashMap<>();
                map.put(String.valueOf(value), value);
                return map;
            }));
            // 图表导入
            rule.addRule("U", "imgHandler", "图片宝", new ImgHandler(true));
            rule.addRule("U", "shortHandler", "short宝", new ShortHandler(true));
            rule.addRule("Y", "stringHandler", "字符串宝", new StringHandler(true));
        };
        PoiResult<DemoData> list = ExcelUtils.readSheet(
                "src/test/java/excel/imports/import.xls",
                DemoData.class, columnRule, 0, 3, 1, (row, rowNum) -> {
                    // 其他逻辑处理,如转换,判断等
                    System.out.println("当前行数据为:" + row);
                });
        if (list.isSuccess()) {
            // 导入没有错误，打印数据
            System.out.println(list.getData().size());
        } else {
            // 导入有错误，打印输出错误
            System.out.println(list.getMessage());
            // 有错误依然可以打印导入的数据
            for (DemoData datum : list.getData()) {

            }
        }
    }
}
```

* 输出结果

```xhtml
[G4]long宝格式不正确
[G6]long宝格式不正确
[G7]long宝格式不正确
[C8]名字格式不正确    [I8]日期宝格式不正确    [K8]double宝格式不正确    [M8]float宝格式不正确    [O8]integer宝格式不正确    [G8]long宝格式不正确    [U8]short宝格式不正确
[C9]名字不能为空
[C10]名字不能为空    
```

## 五. 便捷快速读Excel

* 示例图

![image](https://user-images.githubusercontent.com/29246805/146621587-f9a0f779-84a0-463d-885b-a0d3a16afb3c.png)


* 代码示例

```java
public static void readSheet(){
    List<Map<String, Object>>lists=ExcelUtils.readSheet("src/test/java/excel/export/readExport_OUT.xlsx",0,3,1);
    for(Map<String, Object> list:lists){
        System.out.println(list);
    }
}
```

* 输出

```xhtml
{A=2.0, B=中青旅1, C=, D=华东长三角, E=陕西省, F=保定市, G=张三, H=李四, I=239.0, J=0.33438526939871205, K=Fri Dec 17 17:38:17 CST 2021}
{A=3.0, B=中青旅2, C=, D=华东长三角, E=陕西省, F=保定市, G=张三, H=李四, I=734.0, J=0.046917323921557674, K=Fri Dec 17 17:38:17 CST 2021}
{A=4.0, B=中青旅3, C=, D=华东长三角, E=陕西省, F=保定市, G=张三, H=李四, I=235.0, J=0.8870974305330224, K=Fri Dec 17 17:38:17 CST 2021}
{A=5.0, B=中青旅4, C=, D=西安市, E=陕西省, F=保定市, G=张三, H=李四, I=484.0, J=0.3290520137271864, K=Fri Dec 17 17:38:17 CST 2021}
{A=6.0, B=中青旅5, C=, D=华东长三角, E=陕西省, F=保定市, G=张三, H=李四, I=18.0, J=0.28227006573210534, K=Fri Dec 17 17:38:17 CST 2021}
{A=7.0, B=中青旅6, C=, D=华东长三角, E=陕西省, F=保定市, G=张三, H=李四, I=233.0, J=0.43438315036948694, K=Fri Dec 17 17:38:17 CST 2021}
{A=8.0, B=中青旅7, C=, D=华东长三角, E=陕西省, F=保定市, G=张三, H=李四, I=984.0, J=0.11953560849002731, K=Fri Dec 17 17:38:17 CST 2021}
{A=9.0, B=中青旅8, C=, D=华东长三角, E=陕西省, F=保定市, G=张三, H=李四, I=710.0, J=0.35917838863116225, K=Fri Dec 17 17:38:17 CST 2021}
{A=10.0, B=中青旅9, C=, D=华东长三角, E=陕西省, F=保定市, G=张三, H=李四, I=429.0, J=0.2535268079402, K=Fri Dec 17 17:38:17 CST 2021}
```
