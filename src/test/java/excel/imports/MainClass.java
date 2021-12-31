package excel.imports;

import com.github.stupdit1t.excel.ExcelUtils;
import com.github.stupdit1t.excel.common.PoiResult;
import com.github.stupdit1t.excel.handle.*;
import com.github.stupdit1t.excel.handle.rule.AbsSheetVerifyRule;
import excel.imports.data.DemoData;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

public class MainClass {

    public static void main(String[] args) {
        parseSheet();
        readSheet();
    }

    public static void parseSheet() {
        // 1. 导入规则定义
        Consumer<AbsSheetVerifyRule> columnRule = (rule) -> {
            // 表示C列数据提取到字段bigDecimalHandler,字段为BigDecimal类型, 并且列不能为空
            rule.addRule("C", "bigDecimalHandler", "名字", new LongHandler(false));
            rule.addRule("E", "booleanHandler", "布尔宝", new BooleanHandler(true));
            rule.addRule("G", "charHandler", "char宝", new CharHandler(true));
            rule.addRule("I", "dateHandler", "日期宝", new DateHandler(true, "yyyy-MM-dd HH:mm:ss"));
            rule.addRule("K", "doubleHandler", "double宝", new DoubleHandler(true));
            rule.addRule("M", "floatHandler", "float宝", new FloatHandler(true));
            rule.addRule("O", "integerHandler", "integer宝", new IntegerHandler(true));
            rule.addRule("G", "longHandler", "long宝", new LongHandler(true));
            rule.addRule("S", "objectHandler", "对象宝", new ObjectHandler(true, (value) -> {
                Map<String, Object> map = new HashMap<>();
                map.put(String.valueOf(value), value);
                return map;
            }));
            rule.addRule("U", "imgHandler", "图片宝", new ImgHandler(true));
            rule.addRule("U", "shortHandler", "short宝", new ShortHandler(true));
            rule.addRule("Y", "stringHandler", "字符串宝", new StringHandler(true));
        };
        PoiResult<DemoData> list = ExcelUtils.readSheet(
                "src/test/java/excel/imports/import.xls",
                DemoData.class, columnRule, 0, 3, 1, (row, rowNum) -> {
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

    public static void readSheet() {
        List<Map<String, Object>> lists = ExcelUtils.readSheet("src/test/java/excel/export/readExport_OUT.xlsx", 0, 3, 1);
        for (Map<String, Object> list : lists) {
            System.out.println(list);
        }
    }

}
