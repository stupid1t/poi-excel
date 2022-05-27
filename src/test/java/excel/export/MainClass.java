package excel.export;

import com.github.stupdit1t.excel.core.ExcelUtil;
import com.github.stupdit1t.excel.style.CellPosition;
import excel.export.data.ClassRoom;
import excel.export.data.Parent;
import excel.export.data.ProjectEvaluate;
import excel.export.data.Student;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.io.*;
import java.util.*;

public class MainClass {


    /**
     * 单sheet数据
     */
    private static List<ProjectEvaluate> data = new ArrayList<>();

    /**
     * map型数据
     */
    private static List<Map<String, Object>> mapData = new ArrayList<>();

    /**
     * 复杂对象数据
     */
    private static List<Student> complexData = new ArrayList<>();


    /**
     * 多sheet数据
     */
    private static List<List<?>> moreSheetData = new ArrayList<>();


    static {

        // 1.单sheet数据填充
        for (int i = 0; i < 10; i++) {
            ProjectEvaluate obj = new ProjectEvaluate();
            obj.setProjectName("中青旅" + i);
            obj.setAreaName("华东长三角");
            obj.setProvince("陕西省");
            obj.setCity("保定市");
            obj.setPeople("张三");
            obj.setLeader("李四");
            obj.setScount((int) (Math.random() * 1000));
            obj.setAvg(Math.random());
            obj.setCreateTime(new Date());
            obj.setImg(imageParseBytes(new File("src/test/java/excel/export/data/1.png")));
            data.add(obj);
        }
        // 2.map型数据填充
        for (int i = 0; i < 15; i++) {
            Map<String, Object> obj = new HashMap<>();
            obj.put("name", "张三" + i);
            obj.put("age", 5 + i);
            mapData.add(obj);
        }
        // 3.复杂对象数据
        for (int i = 0; i < 5; i++) {
            // 學生
            Student stu = new Student();
            // 學生所在的班級，用對象
            stu.setClassRoom(new ClassRoom("六班"));
            // 學生的更多信息，用map
            Map<String, Object> moreInfo = new HashMap<>();
            moreInfo.put("parent", new Parent("張無忌"));
            stu.setMoreInfo(moreInfo);
            stu.setName("张三");
            complexData.add(stu);
        }
        // 4.多sheet数据填充
        moreSheetData.add(data);
        moreSheetData.add(mapData);
        moreSheetData.add(complexData);
    }

    public static void main(String[] args) throws IOException {

        long s = System.currentTimeMillis();
        // 简单导出
        simpleExport();
        complexExport();
      /*  simpleExport2();
        complexExport();
        complexExport2();
        mapExport();
        templateExport();
        mulSheet();
        readExport();*/
        System.out.println("耗时:" + (System.currentTimeMillis() - s));

    }

    /**
     * 简单导出
     *
     * @throws Exception
     */
    public static void simpleExport() {
        ExcelUtil.opsExport()
                .password("123456")
                .opsSheet(data)
                .opsHeader().simple().texts("项目名称", "项目图", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间").done()
                .opsColumn().fields("projectName", "img", "areaName", "province", "city", "people", "leader", "scount", "avg", "createTime").done()
                .done()
                .opsSheet(mapData)
                .opsHeader().simple().texts("姓名", "年龄").done()
                .opsColumn().fields("name", "age").done()
                .done()
                .opsSheet(complexData)
                .opsHeader().simple().texts("姓名", "班级").done()
                .opsColumn().fields("name", "classRoom.name").done()
                .done()
                .out("src/test/java/excel/export/simpleExport.xlsx")
        ;
    }


    /**
     * 简单导出
     *
     * @throws Exception
     */
    public static void complexExport() {
        ExcelUtil.opsExport()
                .opsSheet(data)
                .autoNum(true)
                .sheetName("一号sheet")
                .height(CellPosition.TITLE, 1200)
                .height(CellPosition.HEADER, 800)
                .height(CellPosition.FOOTER, 800)
                .opsHeader()
                .complex()
                .text("项目资源统计", "1,1,A,K")
                .text("序号", "2,3,A,A")
                .text("基本信息", "2,2,B,E")
                .text("项目名称", "3,3,B,B")
                .text("所属区域", "3,3,C,C")
                .text("省份", "3,3,D,D")
                .text("市", "3,3,E,E")
                .text("项目所属人", "2,3,F,F")
                .text("市项目领导人", "2,3,G,G")
                .text("分值", "2,2,H,I")
                .text("得分", "3,3,H,H")
                .text("平均分", "3,3,I,I")
                .text("项目图片", "2,3,K,K")
                .text("创建时间", "2,3,J,J")
                .done()
                .opsColumn()
                .fields("projectName", "areaName", "province", "city", "people", "leader", "scount", "avg", "img")
                .field("createTime")
                .align(HorizontalAlignment.LEFT)
                .valign(VerticalAlignment.TOP)
                .height(2000)
                .width(2000)
                .backColor(IndexedColors.BLUE)
                .comment("hhh")
                .datePattern("yyyy-MM-dd")
                .done()
                .done()
                .opsFooter()
                .textIndex("合计:", new Integer[]{0, 1, 0, 2})
                .textIndex("=SUM(H4:H13)", new Integer[]{0, 1, 3, 10})
                .done()
                .done()
                .out("src/test/java/excel/export/complexExport.xlsx");
    }

    /**
     * 将文件转换为byte数组，作为图片数据导入
     *
     * @param file
     * @return byte[]
     */
    public static byte[] imageParseBytes(File file) {
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return imageParseBytes(fileInputStream);
    }

    /**
     * 将流转换为byte数组，作为图片数据导入
     *
     * @param fis
     * @return byte[]
     */
    public static byte[] imageParseBytes(InputStream fis) {
        byte[] buffer = null;
        ByteArrayOutputStream bos = null;
        try {
            bos = new ByteArrayOutputStream(1024);
            byte[] b = new byte[1024];
            int n;
            while ((n = fis.read(b)) != -1) {
                bos.write(b, 0, n);
            }
            buffer = bos.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                fis.close();
                bos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return buffer;
    }

}
