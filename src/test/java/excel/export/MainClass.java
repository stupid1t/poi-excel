package excel.export;

import com.github.stupdit1t.excel.Column;
import com.github.stupdit1t.excel.ExcelUtils;
import com.github.stupdit1t.excel.ExportRules;
import com.github.stupdit1t.excel.style.CellPosition;
import com.github.stupdit1t.excel.style.ICellStyle;
import excel.export.data.ClassRoom;
import excel.export.data.Parent;
import excel.export.data.ProjectEvaluate;
import excel.export.data.Student;
import org.apache.poi.ss.usermodel.*;

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
            obj.setImg(ImageParseBytes(new File("src/test/java/excel/export/data/1.png")));
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
        simpleExport2();
        complexExport();
        complexExport2();
        mapExport();
        templateExport();
        mulSheet();
        readExport();
        System.out.println("耗时:" + (System.currentTimeMillis() - s));

    }

    /**
     * 简单导出
     *
     * @throws Exception
     */
    public static void simpleExport() {
        // 1.导出的header标题设置
        String[] headers = {"项目名称", "项目图", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间"};
        // 2.导出header对应的字段设置
        Column[] columns = {
                Column.field("projectName"),
                Column.field("img"),
                Column.field("areaName"),
                Column.field("province"),
                Column.field("city").width(3),
                Column.field("people"),
                Column.field("leader"),
                Column.field("scount"),
                Column.field("avg"),
                Column.field("createTime").datePattern("yyyy-MM-dd")
        };
        // 3.执行导出
        ExcelUtils.export("src/test/java/excel/export/simpleExport.xlsx", data, ExportRules.simpleRule(columns, headers));
    }

    /**
     * 简单导出2
     *
     * @throws Exception
     */
    public static void simpleExport2() {
        // 1.导出的header标题设置
        String[] headers = {"项目名称", "项目图", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间"};
        // 2.导出header对应的字段设置
        Column[] columns = {
                // 不设置宽度自适应
                Column.field("projectName"),
                // 4.9项目图片
                Column.field("img"),
                // 4.1设置此列宽度为10, 添加注释
                Column.field("areaName").width(10).comment("你好吗"),
                // 4.2设置此列下拉框数据
                Column.field("province").dorpDown(new String[]{"陕西省", "山西省", "辽宁省"}),
                // 4.3设置此列水平居右
                Column.field("city").align(HorizontalAlignment.RIGHT),
                // 4.4 设置此列垂直居上
                Column.field("people").valign(VerticalAlignment.TOP),
                // 4.5 设置此列单元格 自定义校验 只能输入文本
                Column.field("leader")
                        .width(4),
                //.verifyCustom("LEN(G4)>2", "名字长度必须大于2位"),
                // 4.6设置此列单元格 整数 数据校验 ，同时设置背景色为棕色
                Column.field("scount")
                        .verifyIntNum("10~2000")
                        .backColor(IndexedColors.BROWN),
                // 4.7设置此列单元格 浮点数 数据校验， 同时设置字体颜色红色
                Column.field("avg").
                        verifyFloatNum("0.0~20.0")
                        .color(IndexedColors.RED),
                // 4.8设置此列单元格 日期 数据校验 ，同时宽度为20、限制用户表格输入、水平居中、垂直居中、背景色、字体颜色
                Column.field("createTime")
                        .datePattern("yyyy-MM-dd")
                        .verifyDate("2000-01-01~2020-12-12")
                        .align(HorizontalAlignment.LEFT)
                        .valign(VerticalAlignment.CENTER)
                        .backColor(IndexedColors.YELLOW)
                        .color(IndexedColors.GOLD),

        };
        // 3.尾部合计行设计
        Map<String, String> footerRules = new HashMap<>();
        footerRules.put("1,1,A,H", "合计");
        footerRules.put("1,1,I,I", String.format("=SUM(I3:I%s)", 2 + data.size()));
        footerRules.put("1,1,J,J", String.format("=AVERAGE(J3:I%s)", 2 + data.size()));
        footerRules.put("1,1,K,K", "作者:625");

        // 4.自定义header样式
        ICellStyle headerStyle = new ICellStyle() {
            @Override
            public CellPosition getPosition() {
                return CellPosition.HEADER;
            }

            @Override
            public void handleStyle(Font font, CellStyle cellStyle) {
                // 加粗
                font.setBold(true);
                // 黑体
                font.setFontName("黑体");
                // 字号12
                font.setFontHeightInPoints((short) 12);
                // 字体红色
                font.setColor(IndexedColors.RED.getIndex());
                // 背绿色
                cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
                // 边框
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
                cellStyle.setBorderTop(BorderStyle.THIN);
                cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
                // 居中
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                // 折行
                cellStyle.setWrapText(true);
            }
        };
        ExportRules exportRules = ExportRules.simpleRule(columns, headers)
                // 大标题
                .title("简单导出")
                // 自动序号
                .autoNum(true)
                // sheet名称
                .sheetName("简单导出")
                // 尾部合计行设计
                .footerRules(footerRules)
                // 导出格式定义
                .xlsx(true)
                // 自定义全局样式
                .globalStyle(headerStyle);
        // 5.执行导出
        ExcelUtils.export("src/test/java/excel/export/simpleExport2.xlsx", data, exportRules, (fieldName, value, row, col) -> {
                    System.out.print("[打印] 字段:" + fieldName);
                    System.out.print(" 字段值:" + value);
                    System.out.print(" 行数据:" + row);
                    System.out.println(" 单元格样式:" + col);
                    // 设置当前单元格值
                    return value;
                }
        );
    }

    /**
     * 复杂导出
     *
     * @throws Exception
     */
    public static void complexExport() {
        // 1.表头设置,可以对应excel设计表头，一看就懂
        HashMap<String, String> headerRules = new HashMap<>();
        headerRules.put("1,1,A,K", "项目资源统计");
        headerRules.put("2,3,A,A", "序号");
        headerRules.put("2,2,B,E", "基本信息");
        headerRules.put("3,3,B,B", "项目名称");
        headerRules.put("3,3,C,C", "所属区域");
        headerRules.put("3,3,D,D", "省份");
        headerRules.put("3,3,E,E", "市");
        headerRules.put("2,3,F,F", "项目所属人");
        headerRules.put("2,3,G,G", "市项目领导人");
        headerRules.put("2,2,H,I", "分值");
        headerRules.put("3,3,H,H", "得分");
        headerRules.put("3,3,I,I", "平均分");
        headerRules.put("2,3,J,J", "创建时间");
        headerRules.put("2,3,K,K", "项目图片");
        // 2.尾部设置，一般可以用来设计合计栏
        HashMap<String, String> footerRules = new HashMap<>();
        footerRules.put("1,2,A,C", "合计:");
        footerRules.put("1,2,D,K", "=SUM(H4:H13)");
        // 3.导出header对应的字段设置
        Column[] column = {
                Column.field("projectName"),
                Column.field("areaName"),
                Column.field("province"),
                Column.field("city"),
                Column.field("people"),
                Column.field("leader"),
                Column.field("scount"),
                Column.field("avg"),
                Column.field("createTime"),
                Column.field("img")

        };
        // 4.执行导出到工作簿
        ExportRules exportRules = ExportRules.complexRule(column, headerRules)
                .footerRules(footerRules)
                .autoNum(true);
        ExcelUtils.export("src/test/java/excel/export/complexExport.xlsx", data, exportRules);

    }

    /**
     * 复杂的对象级联导出
     *
     * @throws Exception
     */
    public static void complexExport2() {
        // 1.导出的header设置
        String[] header = {"學生姓名", "所在班級", "所在學校", "更多父母姓名"};
        // 2.导出header对应的字段设置，列宽设置
        Column[] column = {
                Column.field("name"),
                Column.field("classRoom.name"),
                Column.field("classRoom.school.name"),
                Column.field("moreInfo.parent.age"),
        };
        // 3.执行导出到工作簿
        ExcelUtils.export("src/test/java/excel/export/complexExport2.xlsx", complexData, ExportRules.simpleRule(column, header));
    }

    /**
     * map数据导出
     *
     * @throws Exception
     */
    public static void mapExport() {
        // 1.导出的header设置
        String[] header = {"姓名", "年龄"};
        // 2.导出header对应的字段设置，列宽设置
        Column[] column = {
                Column.field("name"),
                Column.field("age"),
        };
        // 3.执行导出到工作簿
        ExcelUtils.export("src/test/java/excel/export/mapExport.xlsx", mapData, ExportRules.simpleRule(column, header));
    }

    /**
     * 模板导出
     *
     * @throws Exception
     */
    public static void templateExport() {
        // 1.导出的header设置
        String[] header = {"宝宝姓名", "宝宝昵称", "家长姓名", "手机号码", "宝宝生日", "月龄", "宝宝性别", "来源渠道", "市场人员", "咨询顾问", "客服顾问", "分配校区", "备注"};
        // 2.导出header对应的字段设置，列宽设置
        Column[] column = {Column.field("宝宝姓名"), Column.field("宝宝昵称"),
                Column.field("家长姓名"),
                Column.field("手机号码").verifyText("11~11", "请输入11位的手机号码！"),
                Column.field("宝宝生日").datePattern("yyyy-MM-dd").verifyDate("2000-01-01~3000-12-31"),
                Column.field("月龄").width(4).verifyCustom("VALUE(F3:F6000)", "月齡格式：如1年2个月则输入14"),
                Column.field("宝宝性别").dorpDown(new String[]{"男", "女"}),
                Column.field("来源渠道").width(12).dorpDown(new String[]{"品推", "市场"}),
                Column.field("市场人员").width(6).dorpDown(new String[]{"张三", "李四"}),
                Column.field("咨询顾问").width(6).dorpDown(new String[]{"张三", "李四"}),
                Column.field("客服顾问").width(6).dorpDown(new String[]{"大唐", "银泰"}),
                Column.field("分配校区").width(6).dorpDown(new String[]{"大唐", "银泰"}),
                Column.field("备注")
        };
        // 3.执行导出到工作簿
        ExcelUtils.export("src/test/java/excel/export/templateExport.xlsx", Collections.emptyList(), ExportRules.simpleRule(column, header));
    }

    /**
     * 多sheet导出
     *
     * @throws Exception
     */
    public static void mulSheet() {
        // 1.导出的header设置
        Workbook emptyWorkbook = ExcelUtils.createEmptyWorkbook(true);
        // 2.执行导出到工作簿.1.项目数据2.map数据3.复杂对象数据
        for (int i = 0; i < moreSheetData.size(); i++) {
            if (i == 0) {
                List<ProjectEvaluate> data1 = (ArrayList<ProjectEvaluate>) moreSheetData.get(i);
                // 1.导出的header设置
                String[] header = {"项目名称", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间", "项目图片"};
                // 2.导出header对应的字段设置
                Column[] column = {
                        Column.field("projectName"),
                        Column.field("areaName"),
                        Column.field("province"),
                        Column.field("city"),
                        Column.field("people"),
                        Column.field("leader"),
                        Column.field("scount"),
                        Column.field("avg"),
                        Column.field("createTime"),
                        // 项目图片
                        Column.field("img")

                };
                ExcelUtils.fillBook(emptyWorkbook, data1, ExportRules.simpleRule(column, header).title("项目资源统计").sheetName("mysheet1").autoNum(true));
            }
            if (i == 1) {
                List<Map<String, Object>> data2 = (ArrayList<Map<String, Object>>) moreSheetData.get(i);
                // 1.导出的header设置
                String[] header = {"姓名", "年龄"};
                // 2.导出header对应的字段设置，列宽设置
                Column[] column = {
                        Column.field("name"),
                        Column.field("age"),
                };
                ExcelUtils.fillBook(emptyWorkbook, data2, ExportRules.simpleRule(column, header).sheetName("mysheet2"));
            }

            if (i == 2) {
                List<Student> data3 = (ArrayList<Student>) moreSheetData.get(i);
                // 1.导出的header设置
                String[] header = {"學生姓名", "所在班級", "所在學校", "更多父母姓名"};
                // 2.导出header对应的字段设置，列宽设置
                Column[] column = {
                        Column.field("name"),
                        Column.field("classRoom.name"),
                        Column.field("classRoom.school.name"),
                        Column.field("moreInfo.parent.name"),
                };
                // 3.执行导出到工作簿
                ExcelUtils.fillBook(emptyWorkbook, data3, ExportRules.simpleRule(column, header).title("學生基本信息"));
            }

        }
        // 4.写出文件
        try {
            emptyWorkbook.write(new FileOutputStream("src/test/java/excel/export/mulSheet.xlsx"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 读模板替换变量导出
     *
     * @throws Exception
     */
    public static void readExport() {
        Map<String, String> params = new HashMap<>();
        params.put("author", "625");
        params.put("text", "合计");
        params.put("area", "西安市");
        params.put("prov", "陕西省");
        Workbook workbook = ExcelUtils.readExcelWrite("src/test/java/excel/export/readExport.xlsx", params);
        try {
            workbook.write(new FileOutputStream("src/test/java/excel/export/readExport_OUT.xlsx"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 将流转换为byte数组，作为图片数据导入
     *
     * @param fis
     * @return byte[]
     */
    public static byte[] ImageParseBytes(InputStream fis) {
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

    /**
     * 将文件转换为byte数组，作为图片数据导入
     *
     * @param file
     * @return byte[]
     */
    public static byte[] ImageParseBytes(File file) {
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return ImageParseBytes(fileInputStream);
    }

}
