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
                .variable("projectName", "中青旅")
                .variable("buildName", "管材生产")
                .variable("sendDate", "2020-02-02")
                .variable("reciveSb", "张三")
                .variable("phone", "15594980303")
                .variable("address", "陕西省xxxx")
                .variable("company", FileUtils.readFileToByteArray(new File("C:\\Users\\35361\\Desktop\\1.png")))
                .variable("remark", FileUtils.readFileToByteArray(new File("C:\\Users\\35361\\Desktop\\1.png")))
                .replaceTo("src/test/java/excel/replace/replace2.xlsx");
    }
}
