package excel.imports;

import java.io.FileInputStream;
import java.util.Map;

import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import excel.ExcelUtils;
import excel.ImportRspInfo;
import excel.export.ProjectEvaluate;



public class MainClass {

	public static void main(String[] args) throws Exception {
		FileInputStream fileInputStream = new FileInputStream("src\\test\\java\\excel\\imports\\import.xlsx");
		System.in.read();
	}

	public static void parseSheet() throws Exception {
		// 1.获取源文件
		Workbook wb = WorkbookFactory.create(new FileInputStream("src\\test\\java\\excel\\imports\\import.xlsx"));
		// 2.获取sheet0导入
		Sheet sheet = wb.getSheetAt(0);
		// 3.生成VO实体
		ImportRspInfo<ProjectEvaluate> list = ExcelUtils.parseSheet(ProjectEvaluate.class, EvaluateVerifyBuilder.getInstance(), sheet, 3, 2);
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
		Workbook wb = WorkbookFactory.create(new FileInputStream("src\\test\\java\\excel\\imports\\things.xlsx"));
		Sheet sheet = wb.getSheetAt(0);
		//获取excel中的所有图片
		Map<String, PictureData> pictures = ExcelUtils.getSheetPictures(0, sheet, wb);
		System.out.println(pictures.size());
		// parseSheet
		ImportRspInfo<ProjectEvaluate> list = ExcelUtils.parseSheet(ProjectEvaluate.class, EvaluateVerifyBuilder.getInstance(), sheet, 3, 2, (row, rowNum) -> {
			// 处理图片
			String pictrueIndex = "0," + rowNum + ",12";
			PictureData remove = pictures.remove(pictrueIndex);
			if (null != remove) {
				byte[] data = remove.getData();
				System.out.println(data);
			}
		});
		System.out.println(list);
	}
}
