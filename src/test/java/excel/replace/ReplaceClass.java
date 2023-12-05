package excel.replace;

import com.github.stupdit1t.excel.core.ExcelHelper;
import org.apache.commons.io.FileUtils;
import org.junit.Test;

import java.io.File;
import java.io.IOException;

public class ReplaceClass {

    @Test
    public void parseMap1() throws IOException {
        ExcelHelper.opsReplace()
                .from("src/test/java/excel/replace/replace.xlsx")
                .var("projectName", "中青旅")
                .var("buildName", "管材生产")
                .var("sendDate", "2020-02-02")
                .var("reciveSb", "张三")
                .var("phone", "15594980303")
                .var("address", "陕西省xxxx")
                .var("company", FileUtils.readFileToByteArray(new File("C:\\Users\\35361\\Desktop\\1.png")))
                .var("remark", FileUtils.readFileToByteArray(new File("C:\\Users\\35361\\Desktop\\1.png")))
                .replaceTo("src/test/java/excel/replace/replace2.xlsx");
    }
}
