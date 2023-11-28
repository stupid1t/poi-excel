package excel.parse;

import com.github.stupdit1t.excel.common.Col;
import com.github.stupdit1t.excel.common.PoiResult;
import com.github.stupdit1t.excel.core.ExcelHelper;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.util.HashMap;

public class ParseMapTest {

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
        name.set("快速转map，自动映射列");
        PoiResult<HashMap> parse = ExcelHelper.opsParse(HashMap.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                .opsColumn(true).done()
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }

    @Test
    public void parseMap2() {
        name.set("快速转map，自动映射列，指定列替换");
        PoiResult<HashMap> parse = ExcelHelper.opsParse(HashMap.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                .opsColumn(true)
                    .field(Col.H,"H列保留2位数").type(double.class).scale(2)
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
    public void parseMap3() {
        name.set("快速转map，不自动映射，提取指定列");
        PoiResult<HashMap> parse = ExcelHelper.opsParse(HashMap.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                .opsColumn()
                    .field(Col.A, "name")
                    .field(Col.H,"score")
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
    public void parseMap4() {
        name.set("快速转map，提取指定列，校验");
        PoiResult<HashMap> parse = ExcelHelper.opsParse(HashMap.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 1)
                .opsColumn()
                .field(Col.A, "name").trim().notNull().defaultValue("张三").regex("中青旅\\d{1}")
                .field(Col.H, "score").notNull().scale(2)
                .done()
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }
}


