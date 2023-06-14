package excel.parse;

import com.github.stupdit1t.excel.common.ErrorMessage;
import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.common.PoiResult;
import com.github.stupdit1t.excel.core.ExcelHelper;
import excel.parse.data.ProjectEvaluate;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ParseClass {

    ThreadLocal<Long> time = new ThreadLocal<>();

    ThreadLocal<String> name = new ThreadLocal<>();

    @Before
    public void before() {
        time.set(System.currentTimeMillis());
    }

    @After
    public void after() {
        long diff = System.currentTimeMillis() - time.get();
        System.out.println("[ " + name.get() + " ] 耗时: " + diff);
        time.remove();
        name.remove();
    }

    @Test
    public void parseMap1() {
        name.set("parseMap1");
        PoiResult<HashMap> parse = ExcelHelper.opsParse(HashMap.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                .parse();
        if (!parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }

    @Test
    public void parseMapBig() {
        name.set("parseMapBig");
        ExcelHelper.opsParse(HashMap.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 1)
                .parsePart(1000, (result) -> {
                    if (!result.hasError()) {
                        // 输出验证不通过的信息
                        System.out.println(result.getErrorInfoString());
                    }

                    // 打印解析的数据
                    if (result.hasData()) {
                        result.getData().forEach(System.out::println);
                    }
                });
    }

    @Test
    public void parseBean() {
        name.set("parseBean");
        Map<String, String> mapping = new HashMap<String, String>();
        mapping.put("西安", "1");
        mapping.put("北京", "2");

        PoiResult<ProjectEvaluate> result = ExcelHelper.opsParse(ProjectEvaluate.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                // 自定义列映射
                .opsColumn()
                // 强制输入字符串, 且不能为空
                .field("A", "projectName").asByCustom((row, col, val) -> {
                    if ("中青旅0".equals(val)) {
                        throw new PoiException("数据有误!");
                    }
                    if ("中青旅1".equals(val)) {
                        System.out.println(1 / 0);
                    }
                    // 重写值
                    return val + "-Hello";
                }).notNull().done()
                // img类型. 导入图片必须这样写, 且字段为byte[]
                .field("B", "img").asImg().done()
                .field("C", "areaName").asString().done()
                .field("D", "province").asString().done()
                .field("E", "city").asString().map(mapping::get).done()
                // 不能为空
                .field("F", "people").done()
                // 不能为空
                .field("G", "leader").asString().defaultValue("巨无霸").done()
                // 必须是数字
                .field("H", "scount").asString().done()
                .field("I", "avg").asString().notNull().done()
                .field("J", "createTime").asDate().pattern("yyyy/MM/dd").trim().done()
                .done()
                .map((row, index) -> {
                    // 行回调, 可以在这里改数据
                    System.out.println("当前是第:" + index + " 数据是: " + row);
                    if ("中青旅2-Hello".equals(row.getProjectName())) {
                        throw new NullPointerException();
                    }
                })
                .parse();

        if (result.hasError()) {
            System.out.println("===============单元格错误=================");
            String errorInfoString = result.getErrorInfoString();
            System.out.println(errorInfoString);
            System.out.println("===============数据行错误=================");
            String errorInfoLineString = result.getErrorInfoLineString();
            System.out.println(errorInfoLineString);
            // 获取原始的异常信息
            List<ErrorMessage> error = result.getError();
            // 获取原始的单元格错误
            List<String> errorInfo = result.getErrorInfo();
            // 获取原始的数据行错误
            List<String> errorInfoLine = result.getErrorInfoLine();
        }
    }

    @Test
    public void parseMap3() {
        name.set("parseMap3");
        Map<String, String> cityMapping = new HashMap<String, String>();
        cityMapping.put("西安", "1");
        cityMapping.put("北京", "2");
        PoiResult<HashMap> parse = ExcelHelper.opsParse(HashMap.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                // 自定义列映射
                .opsColumn()
                    // 强制输入字符串, 且不能为空
                    .field("A", "projectName").asString().notNull().done()
                    // img类型. 导入图片必须这样写, 且字段为byte[]
                    .field("B", "img").asImg().done()
                    // 必须为数字
                    .field("C", "areaName").asString().pattern("\\d+").done()
                    .field("D", "province").asFloat().map((val)->{
                        if (val > 100){
                            // 会被收集异常
                            throw PoiException.error("分数太大了!");
                        }
                        return val;
                    }).done()
                    // 城市，去两边空格，并映射字典。
                    .field("E", "city").asString().defaultValue("未知").trim().map(cityMapping::get).done()
                    // 不能为空
                    .field("F", "people").asBigDecimal().notNull().done()
                    // 不能为空
                    .field("G", "leader").asString().notNull().done()
                    // 必须是数字
                    .field("H", "scount").asLong().done()
                    .field("I", "avg").asByCustom((row, col, value) ->{
                        if(value != null){
                            // 自定义处理数据转换验证等操作
                            return 1;
                        }
                        return value;
                    }).done()
                    .field("J", "createTime").asDate().pattern("yyyy-MM-dd").done()
                    .done()
                // 行级别钩子
                .map((row, index) -> {
                    // 行回调, 可以在这里改数据
                    System.out.println("当前是第:" + index + " 数据是: " + row);
                    // 也可以验证数据
                    if (row.get("leader") == null){
                        throw PoiException.error("leader不能为空");
                    }
                })
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }

        // 存在解析通过的数据，打印数据
        if (parse.hasData()) {
            parse.getData().forEach(System.out::println);
        }

    }
}


