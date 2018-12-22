package excel.imports;

import java.io.FileInputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import excel.ExcelUtils;
import excel.ImportRspInfo;
import excel.export.ProjectEvaluate;



public class MainClass {

	public static void main(String[] args) {
		try {
			test();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	public static void parseSheet() throws Exception {
		// 1.获取源文件
		Workbook wb = WorkbookFactory.create(new FileInputStream("src\\test\\java\\excel\\imports\\import.xlsx"));
		// 2.获取sheet0导入
		Sheet sheet = wb.getSheetAt(0);
		// 3.生成VO实体
		ImportRspInfo<ProjectEvaluate> list = ExcelUtils.parseSheet(ProjectEvaluate.class, ProjectVerifyBuilder.getInstance(), sheet, 3, 2);
		if (list.isSuccess()) {
			// 导入没有错误，打印数据
			System.out.println(list.getData());
		} else {
			// 导入有错误，打印输出错误
			System.out.println(list.getMessage());
		}
	}

	/**
	 * 解析excel带回调函数：做一些额外字段填充
	 * 
	 * @throws Exception
	 */
	public static void parseSheetWithCallback() throws Exception {
		Workbook wb = WorkbookFactory.create(new FileInputStream("src\\test\\java\\excel\\imports\\import.xlsx"));
		// parseSheet
		ImportRspInfo<ProjectEvaluate> list = ExcelUtils.parseSheet(ProjectEvaluate.class, ProjectVerifyBuilder.getInstance(), wb.getSheetAt(0), 3, 2, (row, rowNum) -> {
			// 其他逻辑处理
			System.out.println("当前行数据为:"+row);
		});
		if (list.isSuccess()) {
			// 导入没有错误，打印数据
			System.out.println(list.getData());
			//打印图片byte数组长度
			byte[] img = list.getData().get(0).getImg();
			System.out.println(img);
		} else {
			// 导入有错误，打印输出错误
			System.out.println(list.getMessage());
		}
	}

	public static void test() throws Exception {
		Workbook wb = WorkbookFactory.create(new FileInputStream("src\\test\\java\\excel\\imports\\things.xlsx"));
		ImportRspInfo<Things> list = ExcelUtils.parseSheet(Things.class, ThingsVerifyBuilder.getInstance(), wb.getSheetAt(0), 2, 0);
		if (list.isSuccess()) {
			// 导入没有错误，打印数据
			List<Things> data = list.getData();
			for (Things things : data) {
				// 打印图片byte数组长度
				byte[] img = things.getPictureData();
				System.out.println(img);
			}

		} else {
			// 导入有错误，打印输出错误
			System.out.println(list.getMessage());
		}
	}
}
