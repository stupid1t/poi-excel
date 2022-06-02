package excel.replace;

import com.github.stupdit1t.excel.core.ExcelHelper;
import org.junit.Test;

public class ReplaceClass {

    @Test
    public void parseMap1() {
        ExcelHelper.opsReplace()
                .from("src/test/java/excel/replace/replace.xlsx")
                .variable("projectName", "中青旅")
                .variable("buildName", "管材生产")
                .variable("sendDate", "2020-02-02")
                .variable("reciveSb", "张三")
                .variable("phone", "15594980303")
                .variable("address", "陕西省xxxx")
                .variable("company", "社保局")
                .variable("remark", "李四")
                .replaceTo("src/test/java/excel/replace/replace2.xlsx");
    }
}
