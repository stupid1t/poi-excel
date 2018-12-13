# excel-poi （[git地址](http://https://gitee.com/stupid1t/small_tools)）

导出:1.动态表头+表尾，2.支持List<Map>数据,3.支持图片导出，4.支持复杂对象的导出，5.支持回调处理数据后再导出

导入:1.支持严格的单元格校验,2.支持数据行的图片导入,3支持数据回调处理。

导出03和07都支持，默认为03，使用07在调用导出的最后一个参数为true，具体看以下使用方式。

1. 03速度较快，单sheet最大65535行，体积大
2. 07速度慢，单sheet最大1048576行，体积小

非原创，原创地址http://blog.csdn.net/kuyuyingzi/article/details/51323072
加以修改，支持动态生成标题、脚、以及导出格式化指定字段。

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
// 3.获取excel中的所有图片,备用，回调回填到对于的数据实体
Map<String, PictureData> pictures = ExcelUtils.getPictures(0, sheet, wb);
// 4.生成VO数据
//参数：1.生成VO的class类型;2.校验规则;3.导入的sheet;3.从第几行导入;4.尾部非数据行数量;5.导入每条数据的回调
ImportRspInfo<ProjectEvaluate> list = ExcelUtils.parseSheet(ProjectEvaluate.class, EvaluateVerifyBuilder.getInstance(), sheet, 3, 2, (row, rowNum) -> {
	//1.此处可以完成更多的校验
	if(row.getAreaName() == "中青旅"){
	    throw new POIException("第"+rowNum+"行，区域名字不能为中青旅！");
	}
	
	// 2.处理图片，回填到Vo等操作
	String pictrueIndex = "0," + rowNum + ",12";
	PictureData remove = pictures.remove(pictrueIndex);
	if (null != remove) {
		byte[] data = remove.getData();
		row.setPicData(data);
	}
});
if (list.isSuccess()) {
	// 导入没有错误，打印数据
	System.out.println(JSON.toJSONString(list.getData()));
} else {
	// 导入有错误，打印输出错误
	System.out.println(list.getMessage());
}
```

2.1自定义校验器，导入需要校验字段,必须继承AbstractVerifyBuidler

```java
public class EvaluateVerifyBuilder extends AbstractVerifyBuidler {

	private static EvaluateVerifyBuilder builder = new EvaluateVerifyBuilder();

	public static EvaluateVerifyBuilder getInstance() {
		return builder;
	}

	/**
	 * 定义列校验实体：提取的字段、提取列、校验规则1描述字段名称2是否可为空
	 */
	private EvaluateVerifyBuilder() {
		cellEntitys.add(new CellVerifyEntity("projectName", "B", new StringVerify("项目名称", true)));
		cellEntitys.add(new CellVerifyEntity("areaName", "C", new StringVerify("所属区域", true)));
		cellEntitys.add(new CellVerifyEntity("province", "D", new StringVerify("省份", true)));
		cellEntitys.add(new CellVerifyEntity("city", "E", new StringVerify("市", true)));
		cellEntitys.add(new CellVerifyEntity("statusName", "F", new StringVerify("项目状态", true)));
		cellEntitys.add(new CellVerifyEntity("scount", "G", new StringVerify("总分", true)));
		cellEntitys.add(new CellVerifyEntity("areaInfo", "H", new StringVerify("区位条件", true)));
		cellEntitys.add(new CellVerifyEntity("resourceInfo", "I", new StringVerify("资源禀赋", true)));
		cellEntitys.add(new CellVerifyEntity("manageInfo", "G", new StringVerify("经营现状", true)));
		cellEntitys.add(new CellVerifyEntity("reviewInfo", "K", new StringVerify("考察印象", true)));
		cellEntitys.add(new CellVerifyEntity("teamInfo", "L", new StringVerify("管理团队", true)));
		cellEntitys.add(new CellVerifyEntity("img", "M", new StringVerify("风采", true)));
		cellEntitys.add(new CellVerifyEntity("createTime", "N", new DateTimeVerify("创建时间", "yyyy-MM-dd", true)));
		// 必须调用
		super.init();
	}
}

```


#### 导入示例图
![输入图片说明](https://images.gitee.com/uploads/images/2018/1118/104015_a439ba1a_1215820.png "QQ截图20181118104004.png")

### 导出
1.简单导出
```java
//1.获取导出的数据体
List<ProjectEvaluate> data = new ArrayList<ProjectEvaluate>();
//2.导出标题设置，可为空
String title = "项目资源统计";
//3.导出的hearder设置
String[] hearder = { "项目名称", "所属区域", "省份", "市", "项目状态", "字段A", "字段B", "字段C", "字段D", "字段E", "字段F", "字段G" };
//4.导出hearder对应的字段设置，列宽设置
Object[][] fileds = {
		{ "projectName", POIConstant.AUTO },//1.导出字段;2.列宽设置;3.居左居右，默认居中
		{ "areaName", POIConstant.AUTO },
		{ "province", POIConstant.width(5) },//5个汉字长度
		{ "city", POIConstant.AUTO },
		{ "statusName", POIConstant.AUTO },
		{ "scount", POIConstant.AUTO,POIConstant.RIGHT },//该字段，数字设置居右
		{ "areaScore", POIConstant.AUTO },
		{ "resourceScore", POIConstant.AUTO },
		{ "manageScore", POIConstant.AUTO },
		{ "reviewScore", POIConstant.AUTO },
		{ "teamScore", POIConstant.AUTO },
		{ "createTime", POIConstant.charWidth(10) }
};
//5.执行导出到工作簿
//ExportRules:1.是否序号;2.字段信息;3.标题设置可为空;4.表头设置;5.表尾设置可为空，6.是否导出xlsx的07excel，默认参数可以不写为03的xls
Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(false, fileds, title, hearder, null)，true);//指定07xlsx
Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(false, fileds, title, hearder, null)，false);//指定03xls
Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(false, fileds, title, hearder, null));//默认03xls
//6.写出文件
bean.write(new FileOutputStream("src/test/java/example/exp/export.xlsx"));
```

#### 1导出图
![输入图片说明](https://images.gitee.com/uploads/images/2018/1209/193159_168193e2_1215820.png "1.png")


2.复杂表格导出，带回调处理数据逻辑

```java
//1.获取导出的数据体
List<ProjectEvaluate> data = new ArrayList<ProjectEvaluate>();
//2.表头设置,可以对应excel设计表头，一看就懂
HashMap<String, String> headerRules = new HashMap<>();
headerRules.put("1,1,A,M", "项目资源统计");
headerRules.put("2,3,A,A", "序号");
headerRules.put("2,3,B,B", "项目名称");
headerRules.put("2,3,C,C", "所属区域");
headerRules.put("2,3,D,D", "省份");
headerRules.put("2,3,E,E", "市");
headerRules.put("2,3,F,F", "项目状态");
headerRules.put("2,3,G,G", "总分");
headerRules.put("2,2,H,M", "单项评分");
headerRules.put("3,3,H,H", "区位条件");
headerRules.put("3,3,I,I", "资源禀赋");
headerRules.put("3,3,J,J", "经营现状");
headerRules.put("3,3,K,K", "考察印象");
headerRules.put("3,3,L,L", "管理团队");
headerRules.put("3,3,M,M", "创建时间");
//3.尾部设置，一般可以用来设计合计栏
HashMap<String, String> footerRules = new HashMap<>();
footerRules.put("1,2,A,C", "注释:");
footerRules.put("1,2,D,M", "导出参考代码！");
//4.导出字段设置
Object[][] fileds = {
		{ "projectName", POIConstant.AUTO },//1.导出字段;2.列宽设置;3.居左居右，默认居中
		{ "areaName", POIConstant.AUTO },
		{ "province", POIConstant.width(5) },//5个汉字长度
		{ "city", POIConstant.AUTO },
		{ "statusName", POIConstant.AUTO },
		{ "scount", POIConstant.AUTO,POIConstant.RIGHT },//该字段，数字设置居右
		{ "areaScore", POIConstant.AUTO },
		{ "resourceScore", POIConstant.AUTO },
		{ "manageScore", POIConstant.AUTO },
		{ "reviewScore", POIConstant.AUTO },
		{ "teamScore", POIConstant.AUTO },
		{ "createTime", POIConstant.width(10) }
};
//5.执行导出到工作簿
//ExportRules:1.是否序号;2.字段信息;3.标题设置可为空;4.表头设置;5.表尾设置可为空
Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(false, fileds, title, headerRules, footerRules), (fieldName, value, rows) -> {
	System.out.println("当前导出的数据体为:" + rows);
	System.out.println("当前导出值为:" + value);
	// 创建日期，formatter一下
	if (fieldName.equals("createTime")) {
		value = new SimpleDateFormat(POIConstant.FMTDATE).format((Date) (value));
	}
	// 设置图片
	if (fieldName.equals("img")) {
	// 1.假使实体中存储的是文件路径，将路径变成byte返回value就可以
           value = "src/test/java/excel/export/1.png";
	   value = ExcelUtils.ImageParseBytes(new File((String) value));
	}
	return value;
});
//6.写出文件
bean.write(new FileOutputStream("src/test/java/example/exp/export.xlsx"));
```

#### 2导出示例图
![输入图片说明](https://images.gitee.com/uploads/images/2018/1118/105824_6a64ce18_1215820.png "3.png")


3.复杂的对象级联导出

```java
// 1.获取导出的数据体
List<Student> data = new ArrayList<Student>();
//學生
Student stu = new Student();
//學生所在的班級，用對象
stu.setClassRoom(new ClassRoom("六班"));
//學生的更多信息，用map
Map<String,Object> moreInfo = new HashMap<>();
moreInfo.put("parent", new Parent("張無忌"));
stu.setMoreInfo(moreInfo);
stu.setName("张三");
data.add(stu);
// 2.导出标题设置，可为空
String title = "學生基本信息";
// 3.导出的hearder设置
String[] hearder = { "學生姓名", "所在班級", "所在學校", "更多父母姓名"};
// 4.导出hearder对应的字段设置，列宽设置
Object[][] fileds = { 
		{ "name", POIConstant.AUTO }, // 1.导出字段;2.列宽设置;3.居左居右，默认居中
		{ "classRoom.name", POIConstant.AUTO }, 
		{ "classRoom.school.name", POIConstant.AUTO }, // 5个汉字长度
		{ "moreInfo.parent.name", POIConstant.AUTO } 
};
// 5.执行导出到工作簿
// ExportRules:1.是否序号;2.字段信息;3.标题设置可为空;4.表头设置;5.表尾设置可为空
Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(false, fileds, title, hearder, null));
// 6.写出文件
bean.write(new FileOutputStream("src/test/java/excel/export/export3.xlsx"));
```

#### 3导出图
![输入图片说明](https://images.gitee.com/uploads/images/2018/1209/193615_b483f034_1215820.png "4.png")

4.map对象的简单导出

```java
// 1.获取导出的数据体
List<Map<String,String>> data = new ArrayList<Map<String,String>>();
//學生
Map<String,String> map = new HashMap<>();
map.put("name", "张三");
map.put("classRoomName", "三班");
map.put("school", "世纪中心");
map.put("parent", "张无忌");
data.add(map);
// 2.导出标题设置，可为空
String title = "學生基本信息";
// 3.导出的hearder设置
String[] hearder = { "學生姓名", "所在班級", "所在學校", "更多父母姓名"};
// 4.导出hearder对应的字段设置，列宽设置
Object[][] fileds = { 
		{ "name", POIConstant.AUTO }, // 1.导出字段;2.列宽设置;3.居左居右，默认居中
		{ "classRoomName", POIConstant.AUTO }, 
		{ "school", POIConstant.AUTO }, // 5个汉字长度
		{ "parent", POIConstant.AUTO } 
};
// 5.执行导出到工作簿
// ExportRules:1.是否序号;2.字段信息;3.标题设置可为空;4.表头设置;5.表尾设置可为空
Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(false, fileds, title, hearder, null));
// 6.写出文件
bean.write(new FileOutputStream("src/test/java/excel/export/export4.xlsx"));
```

#### 4导出图
![输入图片说明](https://images.gitee.com/uploads/images/2018/1209/193608_c75b81ee_1215820.png "4.png")


# 经常会更新，随时关注哦 :laughing: 

