package excel.export;

import com.github.stupdit1t.excel.common.PoiWorkbookType;
import com.github.stupdit1t.excel.core.ExcelHelper;
import com.github.stupdit1t.excel.style.CellPosition;
import com.github.stupdit1t.excel.style.ICellStyle;
import excel.export.data.ClassRoom;
import excel.export.data.Parent;
import excel.export.data.ProjectEvaluate;
import excel.export.data.Student;
import org.apache.poi.ss.usermodel.*;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.*;
import java.util.*;

public class ExportClass {


    /**
     * 单sheet数据
     */
    private List<ProjectEvaluate> data = new ArrayList<>();

    /**
     * 单sheet数据bigData
     */
    private List<ProjectEvaluate> bigData = new ArrayList<>();

    /**
     * map型数据
     */
    private List<Map<String, Object>> mapData = new ArrayList<>();

    /**
     * 复杂对象数据
     */
    private List<Student> complexData = new ArrayList<>();


    /**
     * 多sheet数据
     */
    private List<List<?>> moreSheetData = new ArrayList<>();

    ThreadLocal<Long> time = new ThreadLocal<>();

    ThreadLocal<String> name = new ThreadLocal<>();

    {

        // 1.单sheet数据填充
        for (int i = 0; i < 10; i++) {
            ProjectEvaluate obj = new ProjectEvaluate();
            obj.setProjectName("中青旅" + i);
            obj.setAreaName("华东长三角");
            obj.setProvince("陕西省");
            if (i % 3 == 0) {
                obj.setCity("北京");
            } else {
                obj.setCity("西安");
            }
            obj.setPeople("张三");
            obj.setLeader("李四");
            obj.setScount((int) (Math.random() * 1000));
            obj.setAvg(Math.random());
            obj.setCreateTime(new Date());
            obj.setImg(imageParseBytes(new File("src/test/java/excel/export/data/1.png")));
            data.add(obj);
        }
        // 1.单sheet数据填充
        for (int i = 0; i < 100000; i++) {
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
            bigData.add(obj);
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

    /**
     * 简单导出
     *
     * @throws Exception
     */
    @Test
    public void simpleExport() {
        name.set("simpleExport");
        ExcelHelper.opsExport(PoiWorkbookType.XLSX)
                .opsSheet(data)
                .opsHeader().simple().texts("项目名称", "项目图", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间").done()
                .opsColumn().fields("projectName", "img", "areaName", "province", "city", "people", "leader", "scount", "avg", "createTime").done()
                .done()
                .getWorkbook("src/test/java/excel/export/excel/simpleExport.xlsx")
        ;
    }

    /**
     * 简单导出
     *
     * @throws Exception
     */
    @Test
    public void simpleExport2() {
        name.set("simpleExport2");

        // 覆盖title全局默认样式
        ICellStyle titleStyle = new ICellStyle() {
            @Override
            public CellPosition getPosition() {
                return CellPosition.TITLE;
            }

            @Override
            public void handleStyle(Font font, CellStyle cellStyle) {
                font.setFontHeightInPoints((short) 100);
                // 红色字体
                font.setColor(IndexedColors.RED.index);
                // 居左
                cellStyle.setAlignment(HorizontalAlignment.LEFT);
            }
        };

        ExcelHelper.opsExport(PoiWorkbookType.XLS)
                // 全局样式覆盖
                .style(titleStyle)
                // 导出添加密码, 仅支持xls格式, 默认无
                .password("123456")
                // sheet声明
                .opsSheet(data)
                // 自动生成序号, 此功能在复杂表头下, 需要自己定义序号列
                .autoNum(true)
                // 自定义数据行高度, 默认excel正常高度
                .height(CellPosition.CELL, 300)
                // 自定义序号列宽度, 默认2000
                .autoNumColumnWidth(3000)
                // sheet名字
                .sheetName("简单导出")
                // 表头标题声明
                .opsHeader().simple()
                // 大标题声明
                .title("我是大标题")
                // 副标题, 自定义样式
                .text("项目名称", (font, style) -> {
                    // 红色
                    font.setColor(IndexedColors.RED.index);
                    // 居顶
                    style.setVerticalAlignment(VerticalAlignment.TOP);
                })
                // 副标题批量
                .texts("项目图", "所属区域", "省份", "项目所属人", "市", "创建时间", "项目领导人", "得分", "平均分")
                .done()
                // 导出列声明
                .opsColumn()
                // 批量导出字段
                .fields("projectName", "img", "areaName", "province", "people")
                // 个性化导出字段设置
                .field("city")
                // 下拉框
                .dropdown("北京", "西安", "上海", "广州")
                // 行高单独设置
                .height(500)
                // 批注
                .comment("城市选择下拉框内容哦")
                // 宽度设置
                .width(6000)
                // 字段导出回调
                .outHandle((val, row, style) -> {
                    // 如果是北京, 设置背景色为黄色
                    if (val.equals("北京")) {
                        style.setBackColor(IndexedColors.YELLOW);
                        style.setHeight(900);
                        // 属性值自定义
                        return val + "(自定义)";
                    }
                    return val;
                }).done()
                .field("createTime")
                // 日期格式化
                .datePattern("yyyy-MM-dd")
                // 居左
                .align(HorizontalAlignment.LEFT)
                // 居中
                .valign(VerticalAlignment.CENTER)
                // 背景黄色
                .backColor(IndexedColors.YELLOW)
                // 金色字体
                .color(IndexedColors.GOLD).done()
                .fields("leader", "scount", "avg")
                .done()
                // 尾行设计
                .opsFooter()
                // 字符合并
                .text("合计", "1,1,A,H")
                // 公式应用
                .text(String.format("=SUM(J3:J%s)", 2 + data.size()), "1,1,J,J")
                .text(String.format("=AVERAGE(K3:K%s)", 2 + data.size()), "1,1,K,K")
                // 坐标合并
                .textIndex("作者:625", new Integer[]{0, 0, 8, 8})
                .done()
                .done()
                // 执行导出
                .getWorkbook("src/test/java/excel/export/excel/simpleExport2.xls")
        ;
    }


    /**
     * 复杂导出
     *
     * @throws Exception
     */
    @Test
    public void complexExport() {
        name.set("complexExport");
        ExcelHelper.opsExport(PoiWorkbookType.XLSX)
                .opsSheet(data)
                .autoNum(true)
                .opsHeader()
                // 不冻结表头
                .freeze(false)
                // 复杂表头模式
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
                .fields("projectName", "areaName", "province", "city", "people", "leader", "scount", "avg", "img", "createTime")
                .done()
                .opsFooter()
                .textIndex("合计:", new Integer[]{0, 1, 0, 2})
                .textIndex("=SUM(H4:H13)", new Integer[]{0, 1, 3, 10})
                .done()
                .done()
                .getWorkbook("src/test/java/excel/export/excel/complexExport.xlsx");
    }

    /**
     * 对象级联导出
     *
     * @throws Exception
     */
    @Test
    public void complexObject() {
        name.set("complexObject");
        ExcelHelper.opsExport(PoiWorkbookType.XLSX)
                .opsSheet(complexData)
                .opsHeader().simple().texts("學生姓名", "所在班級", "所在學校", "更多父母姓名").done()
                .opsColumn().fields("name", "classRoom.name", "classRoom.school.name", "moreInfo.parent.age").done()
                .done()
                .getWorkbook("src/test/java/excel/export/excel/complexObject.xlsx");
    }

    /**
     * map数据导出
     *
     * @throws Exception
     */
    @Test
    public void mapExport() {
        name.set("mapExport");
        ExcelHelper.opsExport(PoiWorkbookType.XLSX)
                .opsSheet(mapData)
                .opsHeader().simple().texts("姓名", "年龄").done()
                .opsColumn().fields("name", "age").done()
                .done()
                .getWorkbook("src/test/java/excel/export/excel/mapExport.xlsx");
    }

    /**
     * 模板导出
     *
     * @throws Exception
     */
    @Test
    public void templateExport() {
        name.set("templateExport");
        ExcelHelper.opsExport(PoiWorkbookType.XLSX)
                .opsSheet(Collections.emptyList())
                .opsHeader().simple().texts("宝宝姓名", "手机号码", "宝宝生日", "月龄", "宝宝性别", "来源渠道", "备注").done()
                .opsColumn()
                .field("宝宝姓名").done()
                .field("手机号码").verifyText("11~11", "请输入11位的手机号码！").done()
                .field("宝宝生日").datePattern("yyyy-MM-dd").verifyDate("2000-01-01~3000-12-31").done()
                .field("月龄").width(4).verifyCustom("VALUE(F3:F6000)", "月齡格式：如1年2个月则输入14").done()
                .field("宝宝性别").dropdown("男", "女").done()
                .field("来源渠道").width(12).dropdown(new String[]{"品推", "市场"}).done()
                .field("备注").done()
                .done()
                .done()
                .getWorkbook("src/test/java/excel/export/excel/templateExport.xlsx");
    }

    /**
     * 多sheet导出
     *
     * @throws Exception
     */
    @Test
    public void mulSheet() {
        name.set("mulSheet");
        ExcelHelper.opsExport(PoiWorkbookType.XLSX)
                // 多线程导出多sheet, 默认false
                .parallelSheet(true)
                .opsSheet(mapData)
                .sheetName("sheet1")
                .opsHeader().simple().texts("姓名", "年龄").done()
                .opsColumn().fields("name", "age").done()
                .done()
                .opsSheet(complexData)
                .sheetName("sheet2")
                .opsHeader().simple().texts("學生姓名", "所在班級", "所在學校", "更多父母姓名").done()
                .opsColumn().fields("name", "classRoom.name", "classRoom.school.name", "moreInfo.parent.age").done()
                .done()
                .opsSheet(bigData)
                .sheetName("sheet3")
                .opsHeader().simple().texts("项目名称", "项目图", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间").done()
                .opsColumn().fields("projectName", "img", "areaName", "province", "city", "people", "leader", "scount", "avg", "createTime").done()
                .done()
                .getWorkbook("src/test/java/excel/export/excel/mulSheet.xlsx");
    }

    /**
     * 模板导出
     *
     * @throws Exception
     */
    @Test
    public void bigData() {
        name.set("bigData 大数据类型");
        ExcelHelper.opsExport(PoiWorkbookType.BIG_XLSX)
                .opsSheet(bigData)
                .sheetName("1")
                .opsHeader().simple().texts("项目名称", "项目图", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间").done()
                .opsColumn().fields("projectName", "img", "areaName", "province", "city", "people", "leader", "scount", "avg", "createTime").done()
                .done()
                .getWorkbook("src/test/java/excel/export/excel/bigData.xlsx");
    }

    /**
     * 将文件转换为byte数组，作为图片数据导入
     *
     * @param file
     * @return byte[]
     */
    public byte[] imageParseBytes(File file) {
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
    public byte[] imageParseBytes(InputStream fis) {
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
