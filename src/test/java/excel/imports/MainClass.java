package excel.imports;

import com.github.stupdit1t.excel.ExcelUtils;
import com.github.stupdit1t.excel.common.ImportRspInfo;
import excel.export.ProjectEvaluate;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MainClass {

	public static void main(String[] args) {
		try {
			readSheet();
			//readExcelWrite();
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
		ImportRspInfo<ProjectEvaluate> list = ExcelUtils.parseSheet(ProjectEvaluate.class, new ProjectVerifyBuilder(), sheet, 3, 0);
		if (list.isSuccess()) {
			// 导入没有错误，打印数据
			System.out.println(list.getData().size());
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
		Workbook wb = WorkbookFactory.create(new FileInputStream("E:\\self\\git\\poi-excel-github\\src\\test\\java\\excel\\imports\\import.xls"));
		// parseSheet
		ImportRspInfo<ProjectEvaluate> list = ExcelUtils.parseSheet(ProjectEvaluate.class, new ProjectVerifyBuilder(), wb.getSheetAt(0), 3, 2, (row, rowNum) -> {
			// 其他逻辑处理
			System.out.println("当前行数据为:" + row);
		});
		if (list.isSuccess()) {
			// 导入没有错误，打印数据
			System.out.println(list.getData());
			// 打印图片byte数组长度
			byte[] img = list.getData().get(0).getImg();
			System.out.println(img);
		} else {
			// 导入有错误，打印输出错误
			System.out.println(list.getMessage());
		}
	}

	public static void readSheet() throws Exception {

		List<Map<String, Object>> lists = ExcelUtils.readSheet("C:\\Users\\damon.li\\Desktop\\123.xlsx",0, 0, 0);
		System.out.println(lists.get(0).size());
	}


	public static void readExcelWrite() throws Exception {
		Map<String,String> params = new HashMap<>();
		params.put("a","今");
		params.put("b","天");
		params.put("c","好");
		params.put("d","开");
		params.put("e","心");
		Workbook workbook = ExcelUtils.readExcelWrite("C:\\Users\\625\\Desktop\\工作簿.xlsx", params);
		workbook.write(new FileOutputStream("C:\\Users\\625\\Desktop\\工作簿 副本.xlsx"));
		workbook.close();
	}


}
