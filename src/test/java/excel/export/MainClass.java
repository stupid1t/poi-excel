package excel.export;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

import excel.ExcelUtils;
import excel.ExcelUtils.Column;
import excel.ExcelUtils.ExportRules;

public class MainClass {

	public static void main(String[] args) throws IOException {
		try {
			long s = System.currentTimeMillis();
			export1();
			export2();
			export3();
			export4();
			export5();
			System.out.println("耗时:" + (System.currentTimeMillis() - s));
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * 
	 * 简单导出
	 * 
	 * @throws Exception
	 */
	public static void export1() throws Exception {
		// 1.获取导出的数据体
		List<ProjectEvaluate> data = new ArrayList<ProjectEvaluate>();
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
			obj.setImg(ExcelUtils.ImageParseBytes(new File("src/test/java/excel/export/1.png")));
			data.add(obj);
		}
		// 2.导出标题设置，可为空
		String title = "项目资源统计";
		// 3.导出的hearder设置
		String[] hearder = { "序号", "项目名称", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间", "项目图片" };
		// 4.导出hearder对应的字段设置
		Column[] column = { Column.field("projectName"), Column.field("areaName"), Column.field("province"),
				Column.field("city"), Column.field("people"), Column.field("leader"), Column.field("scount"),
				Column.field("avg"), Column.field("createTime"),
				// 项目图片
				Column.field("img")

		};
		// 5.执行导出到工作簿
		// ExportRules:1.是否序号;2.列设置;3.标题设置可为空;4.表头设置;5.表尾设置可为空
		Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(true, column, title, hearder, null));
		// 6.写出文件
		bean.write(new FileOutputStream("src/test/java/excel/export/export1.xls"));
	}

	/**
	 * 复杂导出
	 * 
	 * @throws Exception
	 */
	public static void export2() throws Exception {
		// 1.获取导出的数据体
		List<ProjectEvaluate> data = new ArrayList<ProjectEvaluate>();
		for (int i = 0; i < 50; i++) {
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
			obj.setImg(ExcelUtils.ImageParseBytes(new File("src/test/java/excel/export/1.png")));
			data.add(obj);
		}
		// 2.表头设置,可以对应excel设计表头，一看就懂
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
		// 3.尾部设置，一般可以用来设计合计栏
		HashMap<String, String> footerRules = new HashMap<>();
		footerRules.put("1,2,A,C", "注释:");
		footerRules.put("1,2,D,K", "导出参考代码！");
		// 4.导出hearder对应的字段设置
		Column[] column = {

				Column.field("projectName"),
				// 4.1设置此列宽度为10
				Column.field("areaName").width(10),
				// 4.2设置此列下拉框数据
				Column.field("province").width(5).dorpDown(new String[] { "陕西省", "山西省", "辽宁省" }),
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
		// 5.执行导出到工作簿
		// ExportRules:1.是否序号;2.列设置;3.标题设置可为空;4.表头设置;5.表尾设置可为空
		Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(true, column, headerRules, footerRules),
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
		// 6.写出文件
		bean.write(new FileOutputStream("src/test/java/excel/export/export2.xls"));
	}

	/**
	 * 
	 * 复杂的对象级联导出
	 * 
	 * @throws Exception
	 */
	public static void export3() throws Exception {
		// 1.获取导出的数据体
		List<Student> data = new ArrayList<Student>();
		for (int i = 0; i < 10; i++) {
			// 學生
			Student stu = new Student();
			// 學生所在的班級，用對象
			stu.setClassRoom(new ClassRoom("六班"));
			// 學生的更多信息，用map
			Map<String, Object> moreInfo = new HashMap<>();
			moreInfo.put("parent", new Parent("張無忌"));
			stu.setMoreInfo(moreInfo);
			stu.setName("张三");
			data.add(stu);
		}
		// 2.导出标题设置，可为空
		String title = "學生基本信息";
		// 3.导出的hearder设置
		String[] hearder = { "學生姓名", "所在班級", "所在學校", "更多父母姓名" };
		// 4.导出hearder对应的字段设置，列宽设置
		Column[] column = { Column.field("name"), Column.field("classRoom.name"), Column.field("classRoom.school.name"),
				Column.field("moreInfo.parent.name"), };
		// 5.执行导出到工作簿
		// ExportRules:1.是否序号;2.字段信息;3.标题设置可为空;4.表头设置;5.表尾设置可为空
		Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(false, column, title, hearder, null));
		// 6.写出文件
		bean.write(new FileOutputStream("src/test/java/excel/export/export3.xls"));
	}

	/**
	 * 
	 * 简单导出
	 * 
	 * @throws Exception
	 */
	public static void export4() throws Exception {
		// 1.获取导出的数据体
		List<ProjectEvaluate> data = new ArrayList<ProjectEvaluate>();
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
			obj.setImg(ExcelUtils.ImageParseBytes(new File("src/test/java/excel/export/1.png")));
			data.add(obj);
		}
		// 2.导出标题设置，可为空
		String title = "项目资源统计";
		// 3.导出的hearder设置
		String[] hearder = { "序号", "项目名称", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间", "项目图片" };
		// 4.导出hearder对应的字段设置
		Column[] column = { Column.field("projectName"), Column.field("areaName"), Column.field("province"),
				Column.field("city"), Column.field("people"), Column.field("leader"), Column.field("scount"),
				Column.field("avg"), Column.field("createTime"),
				// 项目图片
				Column.field("img")

		};
		// 5.执行导出到工作簿
		// ExportRules:1.是否序号;2.列设置;3.标题设置可为空;4.表头设置;5.表尾设置可为空
		Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(true, column, title, hearder, null));
		// 6.写出文件
		bean.write(new FileOutputStream("src/test/java/excel/export/export1.xls"));
	}

	/**
	 * 
	 * 模板导出
	 * 
	 * @throws Exception
	 */
	public static void export5() throws Exception {
		List<String> list = new ArrayList<>();
		for (int i = 0; i < 50; i++) {
			list.add(i + "");
		}
		String[] drop = list.toArray(new String[list.size()]);

		// 1.导出标题设置，可为空
		String title = "客户导入";
		// 2.导出的hearder设置
		String[] hearder = { "宝宝姓名", "宝宝昵称", "家长姓名", "手机号码", "宝宝生日", "月龄", "宝宝性别", "来源渠道", "市场人员", "咨询顾问", "客服顾问",
				"分配校区", "备注" };
		// 3.导出hearder对应的字段设置，列宽设置
		Column[] column = { Column.field("宝宝姓名"), Column.field("宝宝昵称"), Column.field("家长姓名").dorpDown(drop),
				Column.field("手机号码").verifyText("11~11", "请输入11位的手机号码！"),
				Column.field("宝宝生日").verifyDate("2000-01-01~3000-12-31"),
				Column.field("月龄").width(4).verifyCustom("VALUE(F3:F6000)", "月齡格式：如1年2个月则输入14"),
				Column.field("宝宝性别").dorpDown(new String[] { "男", "女" }),
				Column.field("来源渠道").width(12).dorpDown(new String[] { "品推", "市场" }),
				Column.field("市场人员").width(6).dorpDown(new String[] { "张三", "李四" }),
				Column.field("咨询顾问").width(6).dorpDown(new String[] { "张三", "李四" }),
				Column.field("客服顾问").width(6).dorpDown(new String[] { "大唐", "银泰" }),
				Column.field("分配校区").width(6).dorpDown(new String[] { "大唐", "银泰" }), Column.field("备注") };
		// 5.执行导出到工作簿
		Workbook bean = ExcelUtils.createWorkbook(Collections.emptyList(),
				new ExportRules(false, column, title, hearder, null).setXlsx(true));
		// 6.写出文件
		bean.write(new FileOutputStream("src/test/java/excel/export/export5.xlsx"));
	}

}
