package excel.parse;

import com.github.stupdit1t.excel.common.Col;
import com.github.stupdit1t.excel.common.PoiResult;
import com.github.stupdit1t.excel.core.ExcelHelper;
import excel.parse.data.ProjectEvaluate;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.util.HashMap;
import java.util.Map;

public class ParseBeanTest {

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
    public void parseBean1() {
        name.set("自动映射列");
        PoiResult<ProjectEvaluate> parse = ExcelHelper.opsParse(ProjectEvaluate.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 1)
                .opsColumn(true).done()
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }

        // 打印解析的数据
        System.out.println("数据行数:" + parse.getData().size());
        parse.getData().forEach(System.out::println);
    }

    @Test
    public void parseBean2() {
        name.set("自动映射列，指定列替换");
        PoiResult<ProjectEvaluate> parse = ExcelHelper.opsParse(ProjectEvaluate.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 1)
                .opsColumn(true)
                .field(Col.H, ProjectEvaluate::getScore).type(double.class).scale(2)
                .done()
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }

    @Test
    public void parseBean3() {
        name.set("不自动映射，提取指定列");
        PoiResult<ProjectEvaluate> parse = ExcelHelper.opsParse(ProjectEvaluate.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                .opsColumn()
                .field(Col.A, ProjectEvaluate::getProjectName)
                .field(Col.H, ProjectEvaluate::getScore)
                .done()
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }

    @Test
    public void parseBean4() {
        name.set("提取指定列，校验, 转换");
        Map<String, Integer> cityMapping = new HashMap<>();
        cityMapping.put("西安", 1);
        cityMapping.put("北京", 2);

        PoiResult<ProjectEvaluate> parse = ExcelHelper.opsParse(ProjectEvaluate.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 1)
                .opsColumn()
                .field(Col.A, "projectName").trim().notNull().defaultValue("张三").regex("中青旅\\d{2}")
                .field(Col.D, ProjectEvaluate::getProvince)
                // 值映射转换，也可以异常处理校验等
                .field(Col.E, "cityKey").notNull().map(cityMapping::get)
                .done()
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }

    @Test
    public void parseBean5() {
        name.set("大数据分批处理");
        ExcelHelper.opsParse(ProjectEvaluate.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 1)
                .opsColumn(true).done()
                .parsePart(2, (result) -> {
                    if (result.hasError()) {
                        // 输出验证不通过的信息
                        System.out.println(result.getErrorInfoString());
                    }

                    // 打印解析的数据
                    System.out.println("数据行数:" + result.getData().size());
                    result.getData().forEach(System.out::println);
                });

    }
}


