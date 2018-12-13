package excel.export;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

import excel.ExcelUtils;
import excel.ExcelUtils.ExportRules;
import excel.POIConstant;

public class MainClass {

	public static void main(String[] args) throws IOException {
		try {
			long s = System.currentTimeMillis();
			export1();
			// export2();
			// export3();
			// export4();
			System.out.println("耗时:"+(System.currentTimeMillis()-s));
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
			obj.setStatusName("已签署协议");
			obj.setScount("9.55");
			obj.setAreaScore(1.2332);
			obj.setResourceScore(2.4232);
			obj.setManageScore(3.2323);
			obj.setResourceScore(4.0323);
			obj.setReviewScore(5.2323);
			obj.setTeamScore(6.0234);
			obj.setPotentialScore(21232133.0423);
			obj.setCreateTime(new Date());
			data.add(obj);
		}
		// 2.导出标题设置，可为空
		String title = "项目资源统计";
		// 3.导出的hearder设置
		String[] hearder = { "序号", "项目名称", "所属区域", "省份", "市", "项目状态", "字段A", "字段B", "字段C", "字段D", "字段E", "字段F", "字段" };
		// 4.导出hearder对应的字段设置，列宽设置
		Object[][] fileds = { { "projectName", POIConstant.width(1) }, // 1.导出字段;2.列宽设置;3.居左居右，默认居中
				{ "areaName", POIConstant.width(2) }, { "province", POIConstant.width(3) }, // 5个汉字长度
				{ "city", POIConstant.width(4) }, { "statusName", POIConstant.width(5) }, { "scount", POIConstant.width(6), POIConstant.RIGHT }, // 该字段，数字设置居右
				{ "areaScore", POIConstant.width(7) }, { "resourceScore", POIConstant.width(8) }, { "manageScore", POIConstant.width(9) }, { "reviewScore", POIConstant.width(10) },
				{ "teamScore", POIConstant.width(11) }, { "createTime", POIConstant.width(12) } };
		// 5.执行导出到工作簿
		// ExportRules:1.是否序号;2.字段信息;3.标题设置可为空;4.表头设置;5.表尾设置可为空
		Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(true, fileds, title, hearder, null));
		// 6.写出文件
		bean.write(new FileOutputStream("src/test/java/excel/export/export1.xlsx"));
	}

	/**
	 * 复杂导出
	 * 
	 * @throws Exception
	 */
	public static void export2() throws Exception {
		// 1.获取导出的数据体
		List<ProjectEvaluate> data = new ArrayList<ProjectEvaluate>();
		for (int i = 0; i < 10; i++) {
			ProjectEvaluate obj = new ProjectEvaluate();
			obj.setProjectName("中青旅" + i);
			obj.setAreaName("华东长三角");
			obj.setProvince("河北省");
			obj.setCity("保定市");
			obj.setStatusName("已签署协议");
			obj.setScount("9.55");
			obj.setAreaScore(1.2332);
			obj.setResourceScore(2.4232);
			obj.setManageScore(3.2323);
			obj.setResourceScore(4.0323);
			obj.setReviewScore(5.2323);
			obj.setTeamScore(6.0234);
			obj.setPotentialScore(21232133.0423);
			obj.setCreateTime(new Date());
			data.add(obj);
		}
		// 2.表头设置,可以对应excel设计表头，一看就懂
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
		// 3.尾部设置，一般可以用来设计合计栏
		HashMap<String, String> footerRules = new HashMap<>();
		footerRules.put("1,2,A,C", "注释:");
		footerRules.put("1,2,D,M", "导出参考代码！");
		// 4.导出字段设置
		Object[][] fields = { { "projectName", POIConstant.AUTO }, // 1.导出字段;2.列宽设置;3.居左居右，默认居中
				{ "areaName", POIConstant.AUTO }, { "province", POIConstant.width(5) }, // 5个汉字长度
				{ "city", POIConstant.AUTO }, { "statusName", POIConstant.AUTO },
				{ "scount", POIConstant.AUTO, POIConstant.RIGHT }, // 该字段，数字设置居右
				{ "areaScore", POIConstant.AUTO }, { "resourceScore", POIConstant.AUTO },
				{ "manageScore", POIConstant.AUTO }, { "reviewScore", POIConstant.AUTO },
				{ "teamScore", POIConstant.AUTO }, { "createTime", POIConstant.width(10) } };
		// 5.执行导出到工作簿
		// ExportRules:1.是否序号;2.字段信息;3.标题设置可为空;4.表头设置;5.表尾设置可为空

		Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(true, fields, headerRules, footerRules),
				(fieldName, value, rows) -> {
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
		// 6.写出文件
		bean.write(new FileOutputStream("src/test/java/excel/export/export2.xlsx"));
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
		Object[][] fileds = { { "name", POIConstant.AUTO }, // 1.导出字段;2.列宽设置;3.居左居右，默认居中
				{ "classRoom.name", POIConstant.AUTO }, { "classRoom.school.name", POIConstant.AUTO }, // 5个汉字长度
				{ "moreInfo.parent.name", POIConstant.AUTO } };
		// 5.执行导出到工作簿
		// ExportRules:1.是否序号;2.字段信息;3.标题设置可为空;4.表头设置;5.表尾设置可为空
		Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(false, fileds, title, hearder, null));
		// 6.写出文件
		bean.write(new FileOutputStream("src/test/java/excel/export/export3.xlsx"));
	}

	/**
	 * 
	 * map对象的简单导出
	 * 
	 * @throws Exception
	 */
	public static void export4() throws Exception {
		// 1.获取导出的数据体
		List<Map<String, String>> data = new ArrayList<Map<String, String>>();
		for (int i = 0; i < 10; i++) {
			// 學生
			Map<String, String> map = new HashMap<>();
			map.put("name", "张三");
			map.put("classRoomName", "三班");
			map.put("school", "世纪中心");
			map.put("parent", "张无忌");
			data.add(map);
		}
		// 2.导出标题设置，可为空
		String title = "學生基本信息";
		// 3.导出的hearder设置
		String[] hearder = { "學生姓名", "所在班級", "所在學校", "更多父母姓名" };
		// 4.导出hearder对应的字段设置，列宽设置
		Object[][] fileds = { { "name", POIConstant.AUTO }, // 1.导出字段;2.列宽设置;3.居左居右，默认居中
				{ "classRoomName", POIConstant.AUTO }, { "school", POIConstant.AUTO }, // 5个汉字长度
				{ "parent", POIConstant.AUTO } };
		// 5.执行导出到工作簿
		// ExportRules:1.是否序号;2.字段信息;3.标题设置可为空;4.表头设置;5.表尾设置可为空
		Workbook bean = ExcelUtils.createWorkbook(data, new ExportRules(false, fileds, title, hearder, null));
		// 6.写出文件
		bean.write(new FileOutputStream("src/test/java/excel/export/export4.xlsx"));
	}
}
