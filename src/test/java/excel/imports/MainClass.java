package excel.imports;

import com.github.stupdit1t.excel.ExcelUtils;
import com.github.stupdit1t.excel.common.PoiResult;
import com.github.stupdit1t.excel.handle.*;
import com.github.stupdit1t.excel.handle.rule.AbsSheetVerifyRule;
import excel.imports.data.DemoData;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

public class MainClass {

    public static void main(String[] args) {
        try {
            parseSheet();
            readSheet();
            readExcelWrite();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    public static void parseSheet() throws Exception {
        // 3.生成VO实体
        Consumer<AbsSheetVerifyRule> importRule = (rule) -> {
            rule.addRule("C", "bigDecimalHandler", "名字", new BigDecimalHandler(false));
            rule.addRule("E", "booleanHandler", "booleanHandler", new BooleanHandler(true));
            rule.addRule("G", "charHandler", "charHandler", new CharHandler(true));
            rule.addRule("I", "dateHandler", "dateHandler", new DateHandler("yyyy-MM-dd HH:mm:ss", true));
            rule.addRule("K", "doubleHandler", "doubleHandler", new DoubleHandler(true));
            rule.addRule("M", "floatHandler", "floatHandler", new FloatHandler(true));
            rule.addRule("O", "integerHandler", "integerHandler", new IntegerHandler(true));
            rule.addRule("G", "longHandler", "longHandler", new LongHandler(true));
            rule.addRule("S", "objectHandler", "objectHandler", new ObjectHandler(true, (value) -> {
                Map<String, Object> map = new HashMap<>();
                map.put(String.valueOf(value), value);
                return map;
            }));
            rule.addRule("U", "imgHandler", "图片", new ImgHandler(true));
            rule.addRule("U", "shortHandler", "图片", new ShortHandler(true));
            rule.addRule("Y", "stringHandler", "图片", new StringHandler(true));
        };
        PoiResult<DemoData> list = ExcelUtils.readSheet("src/test/java/excel/imports/import.xls", DemoData.class, importRule, 0, 3, 1, (row, rowNum) -> {
            // 其他逻辑处理
            System.out.println("当前行数据为:" + row);
        });
        if (list.isSuccess()) {
            // 导入没有错误，打印数据
            System.out.println(list.getData().size());
        } else {
            // 导入有错误，打印输出错误
            System.out.println(list.getMessage());
            System.out.println("数据输出");
            for (DemoData datum : list.getData()) {
                System.out.println(datum);
            }
        }
    }

    public static void readSheet() throws Exception {
        List<Map<String, Object>> lists = ExcelUtils.readSheet("C:\\Users\\damon.li\\Desktop\\123.xlsx", 0, 0, 0);
        System.out.println(lists.get(0).size());
    }


    public static void readExcelWrite() throws Exception {
        Map<String, String> params = new HashMap<>();
        params.put("a", "今");
        params.put("b", "天");
        params.put("c", "好");
        params.put("d", "开");
        params.put("e", "心");
        Workbook workbook = ExcelUtils.readExcelWrite("C:\\Users\\625\\Desktop\\工作簿.xlsx", params);
        workbook.write(new FileOutputStream("C:\\Users\\625\\Desktop\\工作簿 副本.xlsx"));
        workbook.close();
    }


}
