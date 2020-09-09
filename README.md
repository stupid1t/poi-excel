[![996.icu](https://img.shields.io/badge/link-996.icu-red.svg)](https://996.icu)
[![LICENSE](https://img.shields.io/badge/license-Anti%20996-blue.svg)](https://github.com/996icu/996.ICU/blob/master/LICENSE)
# Poi-Excel


#### maven使用方式
```java
<!-- excel导入导出 -->
<dependency>
    <groupId>com.github.stupdit1t</groupId>
    <artifactId>poi-excel</artifactId>
    <version>1.5</version>
</dependency>
```
### 优势
> 1. 简单快速上手，且满足绝大多数业务场景
> 2. 屏蔽POI细节，学习成本低。
> 3. 未使用注解方式实现，纯编码代码块，去除烦人的各种POJO
> 4. 功能强大，导入支持严格的单元格校验，导出支持复杂表头和尾部设计,以及单元格样式自定义支持

### 导入细节
1. 支持严格的单元格校验,可以定位到单元格坐标校验
2. 支持数据行的图片导入
3. 支持导入过程中,对数据处理添加回调逻辑,满足其他业务场景
4. xls和xlsx都支持导入

### 导出细节
1. 支持傻瓜式定义动态表头+表尾，如列标题的定义和尾部合计行定义
2. 支持Map/复杂对象(可为Map或者对象嵌套,导出的列则可定义为school.manager.name)/模板/图片导出
3. 支持回调逻辑，处理业务数据化再导出，灵活
4. 支持全局或者局部单元格的样式设置,如颜色大小定位背景色等等
5. xls和xlsx都支持导出
6. 支持多sheet合并导出,多个sheet就可以利用多线程加速导出
7. 支持大数据内存导出，防止OOM，选择SXSSFWorkbook


### 选择xls还是xlsx？
1. xls速度较快，单sheet最大65535行，体积大
2. xlsx速度慢，单sheet最大1048576行，体积小


## 主要功能：
### 导入
1.简单的导入:
```java
// 1.获取源文件
Workbook wb = WorkbookFactory.create(new FileInputStream("src\\test\\java\\excel\\imports\\import.xlsx"));
// 2.获取sheet0导入
Sheet sheet = wb.getSheetAt(0);
// 3.生成VO数据
//参数：1.生成VO的class类型;2.校验规则;3.导入的sheet;3.从第几行导入;4.尾部非数据行数量
ImportRspInfo<ProjectEvaluate> list = ExcelUtils.parseSheet(ProjectEvaluate.class, EvaluateVerifyBuilder.getInstance(), sheet, 3, 2);
if (list.isSuccess()) {
	// 导入没有错误，打印数据
	System.out.println(JSON.toJSONString(list.getData()));
} else {
	// 导入有错误，打印输出错误
	System.out.println(list.getMessage());
}
```

2.复杂导入，带图片导入，带回调处理
```java
// 1.获取源文件
Workbook wb = WorkbookFactory.create(new FileInputStream("src\\test\\java\\excel\\imports\\import.xlsx"));
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

3.自定义校验器，导入需要校验字段,必须继承AbstractVerifyBuidler

```java
public class ProjectVerifyBuilder extends AbstractVerifyBuidler {

	private static ProjectVerifyBuilder builder = new ProjectVerifyBuilder();

	public static ProjectVerifyBuilder getInstance() {
		return builder;
	}

	/**
	 * 定义列校验实体：提取的字段、提取列、校验规则
	 */
	private ProjectVerifyBuilder() {
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
		// 必须调用
		super.init();
	}
}

```


#### 导入示例图
![输入图片说明](https://images.gitee.com/uploads/images/2018/1118/104015_a439ba1a_1215820.png "QQ截图20181118104004.png")

### 导出
0.基础数据构建
```java

    /**
     * 单sheet数据
     */
    static List<ProjectEvaluate> sheetData = new ArrayList<>();

    /**
     * map型数据
     */
    static List<Map<String, Object>> mapData = new ArrayList<>();

    /**
     * 复杂对象数据
     */
    static List<Student> complexData = new ArrayList<>();


    /**
     * 多sheet数据
     */
    static List<List<?>> moreSheetData = new ArrayList<>();


    static {

        // 1.单sheet数据填充
        for (int i = 0; i < 10; i++) {
            ProjectEvaluate obj = new ProjectEvaluate();
            obj.setProjectName("中青旅" + i);
            obj.setAreaName("华东长三角");
            obj.setProvince("河北省");
            obj.setCity("保定市");
            obj.setPeople("张三" + i);
            obj.setLeader("李四" + i);
            obj.setScount(50);
            obj.setAvg(60.0);
            obj.setCreateTime(new Date());
            obj.setImg(ImageParseBytes(new File("src/test/java/excel/export/1.png")));
            sheetData.add(obj);
        }
        // 2.map型数据填充
        for (int i = 0; i < 15; i++) {
            Map<String, Object> obj = new HashMap<>();
            obj.put("name", "张三" + i);
            obj.put("age", 5 + i);
            mapData.add(obj);
        }
        // 3.复杂对象数据
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
        // 4.多sheet数据填充
        moreSheetData.add(sheetData);
        moreSheetData.add(mapData);
        moreSheetData.add(complexData);
    }
```
1.简单导出
```java
// 1.获取导出的数据体
 // 1.导出的hearder设置
String[] hearder = {"项目名称", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间", "项目图片"};
// 2.导出hearder对应的字段设置
Column[] column = {Column.field("projectName"), Column.field("areaName"), Column.field("province"),
        Column.field("city"), Column.field("people"), Column.field("leader"), Column.field("scount"),
        Column.field("avg"), Column.field("createTime"),
        // 项目图片
        Column.field("img")

};
// 3.自定义表头title样式
ICellStyle titleStyle = new ICellStyle() {
    @Override
    public CellPosition getPosition() {
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
// 4.执行导出到工作簿
Workbook bean = ExcelUtils.createWorkbook(sheetData, ExportRules.simpleRule(column, hearder).globalStyle(titleStyle).title("项目资源统计").sheetName("mysheet1").autoNum(true), true,
        (feildName, value, t, customStyle) -> {
            //此处指向回调逻辑，可以修改写入excel的值,以及单元格样式，如颜色等
            return value;
        });
// 4.写出文件
bean.write(new FileOutputStream("src/test/java/excel/export/export1.xlsx"));
```

#### 1导出图
![输入图片说明](https://images.gitee.com/uploads/images/2018/1215/161804_3ddf0b6b_1215820.png "1.png")


2.复杂表格导出

```java
// 1.表头设置,可以对应excel设计表头，一看就懂,此处自动生成序号，需要自己设计序号列在第一位
HashMap<String, String> headerRules = new HashMap<>();
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
footerRules.put("1,2,A,C", "注释:");
footerRules.put("1,2,D,K", "导出参考代码！");
// 3.导出hearder对应的字段设置
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
        // 4.9项目图片
        Column.field("img")

};
// 4.执行导出到工作簿
Workbook bean = ExcelUtils.createWorkbook(
        sheetData,
        ExportRules.complexRule(column, headerRules).autoNum(true).footerRules(footerRules).sheetName("mysheet2"),
        true,
        (fieldName, value, row, col) -> {
            if ("projectName".equals(fieldName) && row.getProjectName().equals("中青旅23")) {
                col.align(HorizontalAlignment.LEFT);
                col.valign(VerticalAlignment.CENTER);
                col.height(2);
                col.backColor(IndexedColors.RED);
                col.color(IndexedColors.YELLOW);
            }
            return value;
        });
// 5.写出文件
bean.write(new FileOutputStream("src/test/java/excel/export/export2.xlsx"));
```

#### 2导出图
![输入图片说明](https://images.gitee.com/uploads/images/2018/1215/161814_61f83ff1_1215820.png "2.png")


3.复杂的对象级联导出

```java
 // 1.导出的hearder设置
String[] hearder = {"學生姓名", "所在班級", "所在學校", "更多父母姓名"};
// 2.导出hearder对应的字段设置，列宽设置
Column[] column = {Column.field("name"), Column.field("classRoom.name"), Column.field("classRoom.school.name"),
        Column.field("moreInfo.parent.name"),};
// 3.执行导出到工作簿
Workbook bean = ExcelUtils.createWorkbook(complexData, ExportRules.simpleRule(column, hearder).title("學生基本信息"), true);
// 4.写出文件
bean.write(new FileOutputStream("src/test/java/excel/export/export3.xlsx"));
```

#### 3导出图
![输入图片说明](https://images.gitee.com/uploads/images/2018/1209/193615_b483f034_1215820.png "4.png")

4.map对象的简单导出

```java
// 1.导出的hearder设置
String[] hearder = {"姓名", "年龄"};
// 2.导出hearder对应的字段设置，列宽设置
Column[] column = {Column.field("name"),
        Column.field("age"),
};
// 3.执行导出到工作簿
Workbook bean = ExcelUtils.createWorkbook(mapData, ExportRules.simpleRule(column, hearder), true);
// 4.写出文件
bean.write(new FileOutputStream("src/test/java/excel/export/export4.xlsx"));
```

#### 4导出图
![输入图片说明](https://images.gitee.com/uploads/images/2018/1209/193608_c75b81ee_1215820.png "4.png")

5.模板导出

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
bean.write(new FileOutputStream("src/test/java/excel/export/export5.xlsx"));
```

#### 5导出图
![输入图片说明](https://images.gitee.com/uploads/images/2018/1215/180646_50cc4004_1215820.png "5.png")

6.多sheet合并导出

```java
 // 1.导出的hearder设置
Workbook emptyWorkbook = ExcelUtils.createEmptyWorkbook(true);
// 2.执行导出到工作簿.1.项目数据2.map数据3.复杂对象数据
for (int i = 0; i < moreSheetData.size(); i++) {
    if (i == 0) {
        List<ProjectEvaluate> data1 = (ArrayList<ProjectEvaluate>) moreSheetData.get(i);
        // 1.导出的hearder设置
        String[] hearder = { "项目名称", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间", "项目图片"};
        // 2.导出hearder对应的字段设置
        Column[] column = {Column.field("projectName"), Column.field("areaName"), Column.field("province"),
                Column.field("city"), Column.field("people"), Column.field("leader"), Column.field("scount"),
                Column.field("avg"), Column.field("createTime"),
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
        Column[] column = {Column.field("name"),
                Column.field("age"),
        };
        ExcelUtils.fillBook(emptyWorkbook, data2, ExportRules.simpleRule(column, hearder).sheetName("mysheet2"));
    }

    if (i == 2) {
        List<Student> data3 = (ArrayList<Student>) moreSheetData.get(i);
        // 1.导出的hearder设置
        String[] hearder = {"學生姓名", "所在班級", "所在學校", "更多父母姓名"};
        // 2.导出hearder对应的字段设置，列宽设置
        Column[] column = {Column.field("name"), Column.field("classRoom.name"), Column.field("classRoom.school.name"),
                Column.field("moreInfo.parent.name"),};
        // 3.执行导出到工作簿
        ExcelUtils.fillBook(emptyWorkbook, data3, ExportRules.simpleRule(column, hearder).title("學生基本信息"));
    }

}
// 4.写出文件
emptyWorkbook.write(new FileOutputStream("src/test/java/excel/export/export6.xlsx"));
```
