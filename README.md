# maven使用方式
```java
<!-- excel导入导出 -->
<dependency>
    <groupId>com.github.stupdit1t</groupId>
    <artifactId>poi-excel</artifactId>
    <version>1.8.1</version>
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
##### 1. 支持傻瓜式定义动态表头/表尾

##### 2. 支持单元格公式

##### 3. 支持导出回调逻辑，处理业务数据化再导出

##### 4. 支持全局或者局部单元格的样式自定义
```java
// 1.导出的hearder设置
String[] hearder = { "项目名称", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间", "项目图片"};
// 2.导出hearder对应的字段设置
Column[] column = {Column.field("projectName"), Column.field("areaName"), Column.field("province"),
        Column.field("city"), Column.field("people"), Column.field("leader"), Column.field("scount"),
        Column.field("avg"), Column.field("createTime"),
        // 项目图片
        Column.field("img")

};
// 3.自定义标题样式
ICellStyle titleStyle = new ICellStyle() {
    @Override
    public CellPosition getPosition() {
        // 样式位置
        return CellPosition.TITLE;
    }

    @Override
    public void handleStyle(Font font, CellStyle style) {
        font.setFontHeightInPoints((short) 15);
        font.setColor(IndexedColors.RED.getIndex());
        font.setBold(true);
        // 左右居中
        style.setAlignment(HorizontalAlignment.CENTER);
        // 上下居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(font);
    }
};
// 3.执行导出到工作簿
Workbook bean = ExcelUtils.createWorkbook(sheetData, ExportRules.simpleRule(column, hearder)
        .title("项目资源统计")
        .autoNum(true)
        .sheetName("mysheet1")
        // 应用样式
        .globalStyle(titleStyle)
        ,
        true );
// 4.写出文件
Workbook bigWorkbook = ExcelUtils.createBigWorkbook();
bean.write(new FileOutputStream("src/test/java/excel/export/export1.xlsx"));
```

##### 5. xls和xlsx都支持导出
* 导出示例图
![输入图片说明](https://images.gitee.com/uploads/images/2020/0924/172558_14a0ac1b_1215820.png "导入文件示例图.png")
* 代码示例(1.2.3.4.5功能 )
```java
// 0.准备数据
List<ProjectEvaluate> listData = new ArrayList<>();

// 1.表头设置,可以对应excel设计表头，一看就懂
Map<String, String> headerRules = new HashMap<>();
headerRules.put("1,1,A,K", "项目资源统计");
headerRules.put("2,3,A,A", "序号");
headerRules.put("2,2,B,E", "基本信息");
headerRules.put("3,3,B,B", "项目名称");
headerRules.put("3,3,C,C", "所属区域");
headerRules.put("3,3,D,D", "省份");
headerRules.put("3,3,E,E", "市");
headerRules.put("2,3,F,F", "项目所属人");
headerRules.put("2,3,G,G", "市项目领导人");
headerRules.put("2,2,H,I", "分值");
headerRules.put("3,3,H,H", "得分");
headerRules.put("3,3,I,I", "平均分");
headerRules.put("2,3,J,J", "创建时间");
headerRules.put("2,3,K,K", "项目图片");

// 2.尾部设置，一般可以用来设计合计栏
HashMap<String, String> footerRules = new HashMap<>();
footerRules.put("1,2,A,C", "合计:");
footerRules.put("1,2,D,K", "=SUM(H4:H13)");

// 3.导出hearder对应的属性设置
Column[] column = {
        Column.field("projectName"),
        // 4.1设置此列宽度为10
        Column.field("areaName").width(10),
        // 4.2设置此列下拉框数据
        Column.field("province").width(5).dorpDown(new String[]{"陕西省", "山西省", "辽宁省"}),
        // 4.3设置此列水平居右
        Column.field("city").align(HorizontalAlignment.RIGHT),
        // 4.4 设置此列垂直居上
        Column.field("people").valign(VerticalAlignment.TOP),
        // 4.5 设置此列单元格 自定义校验 只能输入文本
        Column.field("leader").width(4).verifyCustom("VALUE(F3:F500)", "我是提示"),
        // 4.6设置此列单元格 整数 数据校验 ，同时设置背景色为棕色
        Column.field("scount").verifyIntNum("10~20").backColor(IndexedColors.BROWN),
        // 4.7设置此列单元格 浮点数 数据校验， 同时设置字体颜色红色
        Column.field("avg").verifyFloatNum("10.0~20.0").color(IndexedColors.RED),
        // 4.8设置此列单元格 日期 数据校验 ，同时宽度为20、限制用户表格输入、水平居中、垂直居中、背景色、字体颜色
        Column.field("createTime").width(20).verifyDate("2000-01-03 12:35~3000-05-06 23:23")
                .align(HorizontalAlignment.LEFT).valign(VerticalAlignment.CENTER)
                .backColor(IndexedColors.YELLOW).color(IndexedColors.GOLD),
        // 4.9项目图片, 只有该列是byte[], 自动变为图片
        Column.field("img")

};
// 4.导出规则定义
ExportRules exportRules = ExportRules.complexRule(column, headerRules)
        // 自动生成序号, 需要定义序号列
        .autoNum(true)
        // 尾部合计行定义
        .footerRules(footerRules)
        // 全局表格样式定义, 默认就是这个设置, 可自定义, 如果自定义则以自定义为主
        .globalStyle(DefaultCellStyleEnum.TITLE,DefaultCellStyleEnum.HEADER,DefaultCellStyleEnum.CELL)
        // 表格的大标题, ExportRules.simpleRule()简要表头模式才有用, ExportRules.complexRule()复杂规则不在这里定义,如果定义会报错
        .title("中国好项目")
        // 自定义sheet名字
        .sheetName("我的sheet");
// 5.执行导出到工作簿, 依次叔  $集合数据, $导出规则, $是否导出xlsx格式(导出文件名对应好), $导出回调处理
Workbook bean = ExcelUtils.createWorkbook(listData, exportRules, true, new ExportSheetCallback<ProjectEvaluate>() {
    @Override
    // 此处可处理后写入到Excel, 依次是  $字段名  $字段值  $行记录值  $cell单元格样式
    public Object callback(String fieldName, Object value, ProjectEvaluate projectEvaluate, Column customStyle) {
        return value;
    }
});
// 6.写出文件, 如果是web环境则写到servlet里
bean.write(new FileOutputStream("filePath"));
```
##### 6. 支持Map/复杂对象/模板/图片导出
* Map导出
```java
// 1.导出的hearder设置
List<Map<String, Object>> mapData = new ArrayList<>();
Map<String, Object> obj = new HashMap<>();
obj.put("name", "张三");
obj.put("age", 5);
mapData.add(obj);

// 2.导出hearder对应的字段设置，列宽设置
String[] hearder = {"姓名", "年龄"};
Column[] column = {
        Column.field("name"),
        Column.field("age"),
};
// 3.执行导出到工作簿
Workbook bean = ExcelUtils.createWorkbook(mapData, ExportRules.simpleRule(column, hearder), true);
// 4.写出文件
bean.write(new FileOutputStream(filePath);
```
* 复杂对象
```java
// 0.数据准备
List<Student> complexData = new ArrayList<>();
for (int i = 0; i < 20; i++) {
    // 學生
    Student stu = new Student();
    // 學生所在的班級，用對象
    stu.setClassRoom(new ClassRoom("六班"));
    // 學生的更多信息，用map
    Map<String, Object> moreInfo = new HashMap<>();
    moreInfo.put("parent", new Parent("張無忌"));
    stu.setMoreInfo(moreInfo);
    stu.setName("张三");
    complexData.add(stu);
}

// 1.导出的hearder设置
String[] hearder = {"學生姓名", "所在班級", "所在學校", "更多父母姓名"};

// 2.导出hearder对应的字段设置，列宽设置
Column[] column = {
        Column.field("name"), 
        Column.field("classRoom.name"), 
        Column.field("classRoom.school.name"),
        Column.field("moreInfo.parent.age")
};

// 3.执行导出到工作簿
Workbook bean = ExcelUtils.createWorkbook(complexData, ExportRules.simpleRule(column, hearder).title("學生基本信息"), true);

// 4.写出文件
bean.write(new FileOutputStream(filePath));
```
##### 7. 支持多sheet
```java
// 0.准备数据
List<List<?>> moreSheetData = new ArrayList<>();

// 1.创建空的Book
Workbook emptyWorkbook = ExcelUtils.createEmptyWorkbook(true);

// 2.执行导出到工作簿. 1.项目数据  2.map数据  3.复杂对象数据
for (int i = 0; i < moreSheetData.size(); i++) {
    if (i == 0) {
        List<ProjectEvaluate> data1 = (ArrayList<ProjectEvaluate>) moreSheetData.get(i);
        // 1.导出的hearder设置
        String[] hearder = {"项目名称", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间", "项目图片"};
        // 2.导出hearder对应的字段设置
        Column[] column = {
                Column.field("projectName"), 
                Column.field("areaName"), 
                Column.field("province"),
                Column.field("city"),
                Column.field("people"), 
                Column.field("leader"), 
                Column.field("scount"),
                Column.field("avg"), 
                Column.field("createTime"),
                // 项目图片
                Column.field("img")
        };
        ExcelUtils.fillBook(emptyWorkbook, data1, ExportRules.simpleRule(column, hearder).title("项目资源统计").sheetName("mysheet1").autoNum(true));
    }
    if (i == 1) {
        List<Map<String, Object>> data2 = (ArrayList<Map<String, Object>>) moreSheetData.get(i);
        // 1.导出的hearder设置
        String[] hearder = {"姓名", "年龄"};
        // 2.导出hearder对应的字段设置，列宽设置
        Column[] column = {
                Column.field("name"),
                Column.field("age")
        };
        ExcelUtils.fillBook(emptyWorkbook, data2, ExportRules.simpleRule(column, hearder).sheetName("mysheet2"));
    }

    if (i == 2) {
        List<Student> data3 = (ArrayList<Student>) moreSheetData.get(i);
        // 1.导出的hearder设置
        String[] hearder = {"學生姓名", "所在班級", "所在學校", "更多父母姓名"};
        // 2.导出hearder对应的字段设置，列宽设置
        Column[] column = {
                Column.field("name"), 
                Column.field("classRoom.name"), 
                Column.field("classRoom.school.name"),
                Column.field("moreInfo.parent.name")
        };
        // 3.执行导出到工作簿
        ExcelUtils.fillBook(emptyWorkbook, data3, ExportRules.simpleRule(column, hearder).title("學生基本信息"));
    }

}
// 4.写出文件
emptyWorkbook.write(new FileOutputStream(filePath));
```
##### 8. 支持大数据内存导出，防止OOM
```java
// 0.数据准备
List<Map<String, Object>> mapData = new ArrayList<>();
Map<String, Object> obj = new HashMap<>();
obj.put("name", "张三");
obj.put("age", 5);
mapData.add(obj);

// 1.导出的hearder设置
String[] hearder = {"姓名", "年龄"};

// 2.导出hearder对应的字段设置，列宽设置
Column[] column = {
        Column.field("name"),
        Column.field("age"),
};

// 3.执行导出到工作簿
Workbook bigWorkbook = ExcelUtils.createBigWorkbook();
ExcelUtils.fillBook(bigWorkbook, mapData, ExportRules.simpleRule(column, hearder));

// 4.写出文件
bigWorkbook.write(new FileOutputStream(filePath);
```
##### 9. 模板导出
```java
 // 1.导出的hearder设置
String[] hearder = {"宝宝姓名", "宝宝昵称", "家长姓名", "手机号码", "宝宝生日", "月龄", "宝宝性别", "来源渠道", "市场人员", "咨询顾问", "客服顾问",
    "分配校区", "备注"};

// 2.导出hearder对应的字段设置，列宽设置
Column[] column = {Column.field("宝宝姓名"), Column.field("宝宝昵称"), Column.field("家长姓名"),
    Column.field("手机号码").verifyText("11~11", "请输入11位的手机号码！"),
    Column.field("宝宝生日").verifyDate("2000-01-01~3000-12-31"),
    Column.field("月龄").width(4).verifyCustom("VALUE(F3:F6000)", "月齡格式：如1年2个月则输入14"),
    Column.field("宝宝性别").dorpDown(new String[]{"男", "女"}),
    Column.field("来源渠道").width(12).dorpDown(new String[]{"品推", "市场"}),
    Column.field("市场人员").width(6).dorpDown(new String[]{"张三", "李四"}),
    Column.field("咨询顾问").width(6).dorpDown(new String[]{"张三", "李四"}),
    Column.field("客服顾问").width(6).dorpDown(new String[]{"大唐", "银泰"}),
    Column.field("分配校区").width(6).dorpDown(new String[]{"大唐", "银泰"}), Column.field("备注")};

// 3.执行导出到工作簿
Workbook bean = ExcelUtils.createWorkbook(Collections.emptyList(), ExportRules.simpleRule(column, hearder), true);

// 4.写出文件
bean.write(new FileOutputStream(filePath));
```

## 四. 读Excel模板, 替换变量导出
*  变量写在模板中, 形式为${name}, 如图
![输入图片说明](https://images.gitee.com/uploads/images/2020/0924/171712_10923f36_1215820.png "读模板替换.png")
* 代码示例
```java
Map<String,String> params = new HashMap<>();
params.put("a","今");
params.put("b","天");
params.put("c","好");
params.put("d","开");
params.put("e","心");
Workbook workbook = ExcelUtils.readExcelWrite(sourceFilePath, params);
workbook.write(new FileOutputStream(outFilePath));
workbook.close();
```

## 五. 导入
##### 1. 支持严格的单元格校验,可以定位到单元格坐标校验

##### 2. 支持数据行的图片导入

##### 3. 支持导入过程中,对数据处理添加回调逻辑,满足其他业务场景

##### 4. xls和xlsx都支持导入
* 导入文件示例图
![输入图片说明](https://images.gitee.com/uploads/images/2020/0924/171810_345c981a_1215820.png "导入文件示例图.png")

1.先定义导入校验器
```java
public class ProjectVerifyBuilder extends AbstractVerifyBuidler {

	private static ProjectVerifyBuilder builder = new ProjectVerifyBuilder();

	public static ProjectVerifyBuilder getInstance() {
		return builder;
	}

	/**
	 * 定义列校验实体：提取的字段、提取列、校验规则
	 */
    @Override
    protected List<CellVerifyEntity> buildRule() {
        List<CellVerifyEntity> cellEntitys = new ArrayList<>();
		cellEntitys.add(new CellVerifyEntity("projectName", "B", new StringVerify("项目名称", true)));
		cellEntitys.add(new CellVerifyEntity("areaName", "C", new StringVerify("所属区域", true)));
		cellEntitys.add(new CellVerifyEntity("province", "D", new StringVerify("省份", true)));
		cellEntitys.add(new CellVerifyEntity("city", "E", new StringVerify("市", true)));
		cellEntitys.add(new CellVerifyEntity("people", "F", new StringVerify("项目所属人", true)));
		cellEntitys.add(new CellVerifyEntity("leader", "G", new StringVerify("项目领导人", true)));
		cellEntitys.add(new CellVerifyEntity("scount", "H", new IntegerVerify("总分", true)));
		cellEntitys.add(new CellVerifyEntity("avg", "I", new DoubleVerify("历史平均分", true)));
		cellEntitys.add(new CellVerifyEntity("createTime", "J", new DateTimeVerify("创建时间", "yyyy-MM-dd HH:mm", true)));
		cellEntitys.add(new CellVerifyEntity("img", "K", new ImgVerify("图片", false)));
        retrun cellEntitys;
	}
}
```
2.开始导入
```java
// 1.获取源文件
Workbook wb = WorkbookFactory.create(new FileInputStream(filePath));
// 2.获取sheet0导入
Sheet sheet = wb.getSheetAt(0);
// 3.生成VO数据
//参数：1.生成VO的class类型;2.校验规则;3.导入的sheet;3.从第几行导入;4.尾部非数据行数量;5.导入每条数据的回调
ImportRspInfo<ProjectEvaluate> list = ExcelUtils.parseSheet(ProjectEvaluate.class, ProjectVerifyBuilder.getInstance(), sheet, 3, 2, (row, rowNum) -> {
	//1.此处可以完成更多的校验
	if(row.getAreaName() == "中青旅"){
	    throw new POIException("第"+rowNum+"行，区域名字不能为中青旅！");
	}
	//2.图片导入，再ProjectEvaluate定义类型为byte[]的属性就可以，ProjectVerifyBuilder定义ImgVerfiy校验列.就OK了
});
if (list.isSuccess()) {
	// 导入没有错误，打印数据
	System.out.println(JSON.toJSONString(list.getData()));
	//打印图片byte数组长度
	byte[] img = list.getData().get(0).getImg();
	System.out.println(img);
} else {
	// 导入有错误，打印输出错误
	System.out.println(list.getMessage());
}
```
