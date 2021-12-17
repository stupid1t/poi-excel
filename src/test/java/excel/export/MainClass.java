package excel.export;

import com.github.stupdit1t.excel.Column;
import com.github.stupdit1t.excel.ExcelUtils;
import com.github.stupdit1t.excel.common.ExportRules;
import com.github.stupdit1t.excel.style.CellPosition;
import com.github.stupdit1t.excel.style.ICellStyle;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;

public class MainClass {



    /**
     * 单sheet数据
     */
    private static List<ProjectEvaluate> sheetData = new ArrayList<>();

    /**
     * map型数据
     */
    private static List<Map<String, Object>> mapData = new ArrayList<>();

    /**
     * 复杂对象数据
     */
    private  static List<Student> complexData = new ArrayList<>();


    /**
     * 多sheet数据
     */
    private static List<List<?>> moreSheetData = new ArrayList<>();


    static {

        // 1.单sheet数据填充
        for (int i = 0; i < 10; i++) {
            ProjectEvaluate obj = new ProjectEvaluate();
            obj.setProjectName("中青旅" + i);
            obj.setAreaName("华东长三角"+Math.random());
            obj.setProvince("陕西省");
            obj.setCity("保定市"+i);
            obj.setPeople("张三" + i);
            obj.setLeader("李四" + i);
            obj.setScount((int) (Math.random()*1000));
            obj.setAvg(Math.random());
            obj.setCreateTime(new Date());
            obj.setImg(ImageParseBytes(new File("src/test/java/excel/export/1.png")));
            sheetData.add(obj);
        }
        // 2.map型数据填充
        for (int i = 0; i < 15; i++) {
            Map<String, Object> obj = new HashMap<>();
            obj.put("name", "张三" + i);
            obj.put("age", 5 + i);
            mapData.add(obj);
        }
        // 3.复杂对象数据
        for (int i = 0; i < 20; i++) {
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
        moreSheetData.add(sheetData);
        moreSheetData.add(mapData);
        moreSheetData.add(complexData);
    }

    public static void main(String[] args) throws IOException {
        try {
            long s = System.currentTimeMillis();
            export2();
            //export7();
            //export3();
            //export4();
            //export5();
            //export6();
            System.out.println("耗时:" + (System.currentTimeMillis() - s));
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    /**
     * 简单导出
     *
     * @throws Exception
     */
    public static void export1() throws Exception {
        // 1.导出的hearder设置
        String[] hearder = { "项目名称", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间"};
        // 2.导出hearder对应的字段设置
        Column[] column = {Column.field("projectName"), Column.field("areaName").width(30), Column.field("province"),
                Column.field("city"), Column.field("people"), Column.field("leader"), Column.field("scount"),
                Column.field("avg"), Column.field("createTime")
        };
        // 3.执行导出到工作簿
        ICellStyle titleStyle = new ICellStyle() {
            @Override
            public CellPosition getPosition() {
                return CellPosition.CELL;
            }

            @Override
            public void handleStyle(Font font, CellStyle style) {
                font.setFontHeightInPoints((short) 15);
                font.setColor(IndexedColors.RED.getIndex());
                font.setBold(true);
                // 左右居中
                style.setAlignment(HorizontalAlignment.CENTER);
                // 上下居中
                style.setVerticalAlignment(VerticalAlignment.CENTER);
                style.setFont(font);
                //折行显示
                style.setWrapText(true);
            }
        };
        Workbook bigWorkbook = ExcelUtils.createBigWorkbook(500);
        ExcelUtils.fillBook(bigWorkbook,sheetData, ExportRules.simpleRule(column, hearder)
                .title("项目资源统计")
                .autoNum(true)
                .sheetName("mysheet1")
                .globalStyle(titleStyle)
                );
        // 4.写出文件
        bigWorkbook.write(new FileOutputStream("src/test/java/excel/export/export1.xlsx"));
    }

    /**
     * 复杂导出
     *
     * @throws Exception
     */
    public static void export2() throws Exception {
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
        // 3.导出hearder对应的字段设置
        Column[] column = {
                Column.field("projectName"),
                // 4.1设置此列宽度为10
                Column.field("areaName").width(10).comment("你好吗"),
                // 4.2设置此列下拉框数据
                Column.field("province").width(5).dorpDown(new String[]{"陕西省", "山西省", "辽宁省"}),
                // 4.3设置此列水平居右
                Column.field("city").align(HorizontalAlignment.RIGHT),
                // 4.4 设置此列垂直居上
                Column.field("people").valign(VerticalAlignment.TOP),
                // 4.5 设置此列单元格 自定义校验 只能输入文本
                Column.field("leader").width(4).verifyCustom("LEN(G4)>2", "名字长度必须大于2位"),
                // 4.6设置此列单元格 整数 数据校验 ，同时设置背景色为棕色
                Column.field("scount").verifyIntNum("10~2000").backColor(IndexedColors.BROWN),
                // 4.7设置此列单元格 浮点数 数据校验， 同时设置字体颜色红色
                Column.field("avg").verifyFloatNum("0.0~20.0").color(IndexedColors.RED),
                // 4.8设置此列单元格 日期 数据校验 ，同时宽度为20、限制用户表格输入、水平居中、垂直居中、背景色、字体颜色
                Column.field("createTime").width(20).verifyDate("2000-01-03 12:35:12~3000-05-06 23:23:13")
                        .datePattern("yyyy-MM-dd")
                        .align(HorizontalAlignment.LEFT).valign(VerticalAlignment.CENTER)
                        .backColor(IndexedColors.YELLOW).color(IndexedColors.GOLD),
                // 4.9项目图片
                Column.field("img")

        };
        // 4.执行导出到工作簿
        Workbook bean = ExcelUtils.createWorkbook(
                sheetData,
                ExportRules.complexRule(column, headerRules).autoNum(true).footerRules(footerRules).sheetName("mysheet2").xlsx(true),
                (fieldName, value, row, col) -> {
                    if ("projectName".equals(fieldName) && row.getProjectName().equals("中青旅23")) {
                        col.align(HorizontalAlignment.LEFT);
                        col.valign(VerticalAlignment.CENTER);
                        col.height(2);
                        col.backColor(IndexedColors.RED);
                        col.color(IndexedColors.YELLOW);
                    }
                    return value;
                });
        // 5.写出文件
        bean.write(new FileOutputStream("src/test/java/excel/export/export2.xlsx"));
    }

    /**
     * 复杂的对象级联导出
     *
     * @throws Exception
     */
    public static void export3() throws Exception {
        // 1.导出的hearder设置
        String[] hearder = {"學生姓名", "所在班級", "所在學校", "更多父母姓名"};
        // 2.导出hearder对应的字段设置，列宽设置
        Column[] column = {Column.field("name"), Column.field("classRoom.name"), Column.field("classRoom.school.name"),
                Column.field("moreInfo.parent.age"),};
        // 3.执行导出到工作簿
        Workbook bean = ExcelUtils.createWorkbook(complexData, ExportRules.simpleRule(column, hearder).title("學生基本信息"));
        // 4.写出文件
        bean.write(new FileOutputStream("src/test/java/excel/export/export3.xlsx"));
    }

    /**
     * map数据导出
     *
     * @throws Exception
     */
    public static void export4() throws Exception {
        // 1.导出的hearder设置
        String[] hearder = {"姓名", "年龄"};
        // 2.导出hearder对应的字段设置，列宽设置
        Column[] column = {Column.field("name"),
                Column.field("age"),
        };
        // 3.执行导出到工作簿
        Workbook bean = ExcelUtils.createWorkbook(mapData, ExportRules.simpleRule(column, hearder));
        // 4.写出文件
        bean.write(new FileOutputStream("src/test/java/excel/export/export4.xlsx"));
    }

    /**
     * 模板导出
     *
     * @throws Exception
     */
    public static void export5() throws Exception {
        // 1.导出的hearder设置
        String[] hearder = {"宝宝姓名", "宝宝昵称", "家长姓名", "手机号码", "宝宝生日", "月龄", "宝宝性别", "来源渠道", "市场人员", "咨询顾问", "客服顾问",
                "分配校区", "备注"};
        // 2.导出hearder对应的字段设置，列宽设置
        Column[] column = {Column.field("宝宝姓名"), Column.field("宝宝昵称"), Column.field("家长姓名"),
                Column.field("手机号码").verifyText("11~11", "请输入11位的手机号码！"),
                Column.field("宝宝生日").verifyDate("2000-01-01~3000-12-31"),
                Column.field("月龄").width(4).verifyCustom("VALUE(F3:F6000)", "月齡格式：如1年2个月则输入14"),
                Column.field("宝宝性别").dorpDown(new String[]{"男", "女"}),
                Column.field("来源渠道").width(12).dorpDown(new String[]{"品推", "市场"}),
                Column.field("市场人员").width(6).dorpDown(new String[]{"张三", "李四"}),
                Column.field("咨询顾问").width(6).dorpDown(new String[]{"张三", "李四"}),
                Column.field("客服顾问").width(6).dorpDown(new String[]{"大唐", "银泰"}),
                Column.field("分配校区").width(6).dorpDown(new String[]{"大唐", "银泰"}), Column.field("备注")};
        // 3.执行导出到工作簿
        Workbook bean = ExcelUtils.createWorkbook(Collections.emptyList(), ExportRules.simpleRule(column, hearder));
        // 4.写出文件
        bean.write(new FileOutputStream("src/test/java/excel/export/export5.xlsx"));
    }

    /**
     * 多sheet导出,并携带回调
     *
     * @throws Exception
     */
    public static void export6() throws Exception {
        // 1.导出的hearder设置
        Workbook emptyWorkbook = ExcelUtils.createEmptyWorkbook(true);
        // 2.执行导出到工作簿.1.项目数据2.map数据3.复杂对象数据
        for (int i = 0; i < moreSheetData.size(); i++) {
            if (i == 0) {
                List<ProjectEvaluate> data1 = (ArrayList<ProjectEvaluate>) moreSheetData.get(i);
                // 1.导出的hearder设置
                String[] hearder = {"项目名称", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间", "项目图片"};
                // 2.导出hearder对应的字段设置
                Column[] column = {Column.field("projectName"), Column.field("areaName"), Column.field("province"),
                        Column.field("city"), Column.field("people"), Column.field("leader"), Column.field("scount"),
                        Column.field("avg"), Column.field("createTime"),
                        // 项目图片
                        Column.field("img")

                };
                ExcelUtils.fillBook(emptyWorkbook, data1, ExportRules.simpleRule(column, hearder).title("项目资源统计").sheetName("mysheet1").autoNum(true));
            }
            if (i == 1) {
                List<Map<String, Object>> data2 = (ArrayList<Map<String, Object>>) moreSheetData.get(i);
                // 1.导出的hearder设置
                String[] hearder = {"姓名", "年龄"};
                // 2.导出hearder对应的字段设置，列宽设置
                Column[] column = {Column.field("name"),
                        Column.field("age"),
                };
                ExcelUtils.fillBook(emptyWorkbook, data2, ExportRules.simpleRule(column, hearder).sheetName("mysheet2"));
            }

            if (i == 2) {
                List<Student> data3 = (ArrayList<Student>) moreSheetData.get(i);
                // 1.导出的hearder设置
                String[] hearder = {"學生姓名", "所在班級", "所在學校", "更多父母姓名"};
                // 2.导出hearder对应的字段设置，列宽设置
                Column[] column = {Column.field("name"), Column.field("classRoom.name"), Column.field("classRoom.school.name"),
                        Column.field("moreInfo.parent.name"),};
                // 3.执行导出到工作簿
                ExcelUtils.fillBook(emptyWorkbook, data3, ExportRules.simpleRule(column, hearder).title("學生基本信息"));
            }

        }
        // 4.写出文件
        emptyWorkbook.write(new FileOutputStream("src/test/java/excel/export/export6.xlsx"));
    }

    /**
     * 复杂导出
     *
     * @throws Exception
     */
    public static void export7() throws Exception {
        // 1.表头设置,可以对应excel设计表头，一看就懂
        HashMap<String, String> headerRules = new HashMap<>();
        headerRules.put("1,1,A,K", "项目资源统计");
        headerRules.put("2,2,A,K", "序号");
        headerRules.put("3,3,A,K", "序号2");
        headerRules.put("4,4,A,K", "序号2");
        // 3.导出hearder对应的字段设置
        Column[] column = {
                Column.field("projectName"),
                // 4.1设置此列宽度为10
                Column.field("areaName").width(10).comment("你好吗"),
                // 4.2设置此列下拉框数据
                Column.field("province").width(5).dorpDown(new String[]{"陕西省", "山西省", "辽宁省"}),
                // 4.3设置此列水平居右
                Column.field("city").align(HorizontalAlignment.RIGHT),
                // 4.4 设置此列垂直居上
                Column.field("people").valign(VerticalAlignment.TOP),
                // 4.5 设置此列单元格 自定义校验 只能输入文本
                Column.field("leader").width(4).verifyCustom("VALUE(F3:F500)", "我是提示"),
                // 4.6设置此列单元格 整数 数据校验 ，同时设置背景色为棕色
                Column.field("scount").verifyIntNum("10~20").backColor(IndexedColors.BROWN),
                // 4.7设置此列单元格 浮点数 数据校验， 同时设置字体颜色红色
                Column.field("avg").verifyFloatNum("10.0~20.0").color(IndexedColors.RED),
                // 4.8设置此列单元格 日期 数据校验 ，同时宽度为20、限制用户表格输入、水平居中、垂直居中、背景色、字体颜色
                Column.field("createTime").width(20).verifyDate("2000-01-03 12:35~3000-05-06 23:23")
                        .align(HorizontalAlignment.LEFT).valign(VerticalAlignment.CENTER)
                        .backColor(IndexedColors.YELLOW).color(IndexedColors.GOLD),
                // 4.9项目图片
                Column.field("img")

        };
        // 4.执行导出到工作簿
        Workbook bean = ExcelUtils.createWorkbook(
                sheetData,
                ExportRules.complexRule(column, headerRules).autoNum(true).sheetName("mysheet2").xlsx(false),
                (fieldName, value, row, col) -> {
                    if ("projectName".equals(fieldName) && row.getProjectName().equals("中青旅23")) {
                        col.align(HorizontalAlignment.LEFT);
                        col.valign(VerticalAlignment.CENTER);
                        col.height(2);
                        col.backColor(IndexedColors.RED);
                        col.color(IndexedColors.YELLOW);
                    }
                    return value;
                });
        // 5.写出文件
        bean.write(new FileOutputStream("src/test/java/excel/export/export7.xls"));
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
