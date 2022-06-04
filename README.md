# maven使用方式
```xml
<!-- excel导入导出 -->
<dependency>
    <groupId>com.github.stupdit1t</groupId>
    <artifactId>poi-excel</artifactId>
    <version>3.0.2</version>
</dependency>
```

## 一. 项目优势

- 简单快速上手，且满足绝大多数业务场景
- 屏蔽POI细节，学习成本低。
- 未使用注解方式实现，纯编码代码块，去除烦人的各种POJO
- 功能强大，导入支持严格的单元格校验，导出支持公式，复杂表头和尾部设计，以及单元格样式自定义支持
- 支持读取复杂Excel模板,替换变量输出Excel,变量用${}替代

## 二. 更新记录

> 有需求才有进步，这个轮子本身就是从0开始因为需求慢慢叠加起来的。有新需求提出来,我觉得合适会更新. 如有疑问可加群帮解答: 811606008

### v3.0.2 ( 不兼容[1.x.x](README-1.x.md) 和 [2.x.x](README-2.x.md) 版本)

1. 提供ExcelHepler链式构建类, 帮助快捷构建. 本身还是调用ExcelUtil类
2. 优化代码结构和层次
3. 提供更精确的单元格样式控制

## 三. 导出

##### 选择xls还是xlsx？

> xls速度较快，单sheet最大65535行，体积大. xlsx速度慢，单sheet最大1048576行，体积小

##### 1. 简单导出

* 代码示例

```java
class a {

    @Test
    public void simpleExport() {
        // 指定导出XLSX格式
        ExcelHelper.opsExport(PoiWorkbookType.XLSX)
                .opsSheet(data)
                .opsHeader().simple().texts("项目名称", "项目图", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间").done()
                .opsColumn().fields("projectName", "img", "areaName", "province", "city", "people", "leader", "scount", "avg", "createTime").done()
                .done()
                .export("src/test/java/excel/export/excel/simpleExport.xlsx")
        ;
    }
}
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/171567064-3b62d725-96c8-4ee1-b91c-59f1d5f98f56.png)


##### 2. 简单导出 + 自定义属性完整示例

* 代码示例

```java
class a {

    @Test
    public void simpleExport2() {
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

        // 指定导出XLS格式, 只有这个格式支持密码设置
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
                            })
                            .done()
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
                            .color(IndexedColors.GOLD)
                            .done()
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
                        .text("作者:625", 0, 0, 8, 8)
                        .done()
                    .done()
                // 执行导出
                .export("src/test/java/excel/export/excel/simpleExport2.xls")
        ;
    }
}
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/171567183-b3e3b2ca-0bd9-42b5-b521-6167fa9f8279.png)



##### 3. 复杂表头导出

* 代码示例

```java
class a {

    @Test
    public void complexExport() {
        ExcelHelper.opsExport(PoiWorkbookType.XLSX)
                .opsSheet(data)
                    .autoNum(true)
                    .opsHeader()
                        // 不冻结表头
                        .freeze(false)
                        // 复杂表头模式, 支持三种合并方式, 1数字坐标 2字母坐标 3Excel坐标
                        .complex()
                            // excel坐标
                            .text("项目资源统计", "A1:K1")
                            // 字母坐标
                            .text("序号", "2,3,A,A")
                            // 数字坐标
                            .text("基本信息", 1,1,1,4)
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
                            // 尾行合计,D1,K2中的 纵坐标从1开始计算,会自动计算数据行高度!  切记! 切记! 切记!
                            .text("=SUM(H4:H13)", "D1:K2")
                            .text("=SUM(H4:H13)", 0, 1, 3, 10)
                            .done()
                        // 自定义合并单元格
                        .mergeCell("F4:G13")
                        .done()
                    .export("src/test/java/excel/export/excel/complexExport.xlsx");
    }
}
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/172006040-db5b4016-a54e-4816-8585-bf760a4f8e54.png)





##### 4. 复杂的对象级联导出

* 代码示例

```java
class a {

    @Test
    public void complexObject() {
        ExcelHelper.opsExport(PoiWorkbookType.XLSX)
                .opsSheet(complexData)
                .opsHeader().simple().texts("學生姓名", "所在班級", "所在學校", "更多父母姓名").done()
                .opsColumn().fields("name", "classRoom.name", "classRoom.school.name", "moreInfo.parent.age").done()
                .done()
                .export("src/test/java/excel/export/excel/complexObject.xlsx");
    }
}
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/171567981-1690ea1e-7116-40de-82b0-3ae3aa3a0754.png)



##### 4. map数据导出

```java
class a {

    List<Map<String, Object>> mapData = new ArrayList<>();

    @Test
    public void mapExport() {
        ExcelHelper.opsExport(PoiWorkbookType.XLSX)
                .opsSheet(mapData)
                .opsHeader().simple().texts("姓名", "年龄").done()
                .opsColumn().fields("name", "age").done()
                .done()
                .export("src/test/java/excel/export/excel/mapExport.xlsx");
    }
}
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/171568035-6ae3b80f-3453-4e21-b0f6-bc2703bcdf8f.png)



##### 5. 模板导出

* 代码示例

```java
class a {

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
                        .field("来源渠道").width(12).dropdown("品推", "市场").done()
                        .field("备注").done()
                        .done()
                    .done()
                .export("src/test/java/excel/export/excel/templateExport.xlsx");
    }
}
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/171568101-1ab4d733-85be-4b7e-ac88-82098f790d4f.png)



##### 6. 支持多sheet导出

* 代码示例

```java
class a {

    @Test
    public void mulSheet() {
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
                .export("src/test/java/excel/export/excel/mulSheet.xlsx");
    }
}
```
* 导出示例

![image](https://user-images.githubusercontent.com/29246805/171568172-4fd123d5-dc6f-49a8-965a-a7f76751fa46.png)


##### 7. 支持大数据内存导出，防止OOM

* 代码示例

```java

@Test
class a {
    public void bigData() {
        // 声明导出BIG XLSX
        ExcelHelper.opsExport(PoiWorkbookType.BIG_XLSX)
                .opsSheet(bigData)
                .sheetName("1")
                .opsHeader().simple().texts("项目名称", "项目图", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间").done()
                .opsColumn().fields("projectName", "img", "areaName", "province", "city", "people", "leader", "scount", "avg", "createTime").done()
                .done()
                .export("src/test/java/excel/export/excel/bigData.xlsx");
    }
}
```

##### 8. 读模板替换变量导出
![image](https://user-images.githubusercontent.com/29246805/171568236-9960f08c-782d-4457-92e8-a681cbcb06e6.png)

        
* 代码示例

```java
class a {

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
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/171568276-76572937-b483-441c-bc53-10ea3eea0b4d.png)


## 五. 解析导入
##### 1. 支持严格的单元格校验,可以定位到单元格坐标校验
##### 2. 支持数据行的图片导入
##### 3. 支持导入过程中,对数据处理添加回调逻辑,满足其他业务场景
##### 4. xls和xlsx都支持导入

* 导入文件示例图

![image](https://user-images.githubusercontent.com/29246805/171568340-39a10aed-1c00-405c-a7f8-3157ab40f10d.png)



* 数据要转换的Java对象

```java
public class ProjectEvaluate implements Serializable {

    private Long id;

    private Long projectId;

    private Long createUserId;

    private Date createTime;

    private String projectName;

    private String areaName;

    private String province;

    private String city;

    private String statusName;

    private Integer scount;

    private double avg;

    private String people;

    private String leader;

    private byte[] img;
}
```

* 快速解析为Map, 不验证excel内容

```java
public class MainClass {

    @Test
    public void parseMap1() {
        PoiResult<Map> parse = ExcelHelper.opsParse(Map.class)
                .from("src/test/java/excel/export/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                .parse();
        if (!parse.isSuccess()) {
            // 输出验证不通过的信息
            System.out.println(parse.getMessageToString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }
}
```

* 快速解析为Map, 验证excel内容

```java
public class MainClass {

    @Test
    public void parseMap2() {
        PoiResult<Map> parse = ExcelHelper.opsParse(Map.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                // 自定义列映射
                .opsColumn()
                    // 强制输入字符串, 且不能为空
                    .field("A", "projectName", "项目名称").asString().notNull().done()
                    // img类型. 导入图片必须这样写, 且字段为byte[]
                    .field("B", "img", "项目图片").asImg().done()
                    .field("C", "areaName", "所属区域").done()
                    .field("D", "province", "省份").done()
                    .field("E", "city", "市").done()
                    // 不能为空
                    .field("F", "people", "项目所属人").notNull().done()
                    // 不能为空
                    .field("G", "leader", "项目领导人").notNull().done()
                    // 必须是数字
                    .field("H", "scount", "总分").asLong().done()
                    .field("I", "avg", "历史平均分").done()
                    .field("J", "createTime", "创建时间").asDate("yyyy-MM-dd").done()
                    .done()
                .callBack((row, index) -> {
                    // 行回调, 可以在这里改数据
                    System.out.println("当前是第:" + index + " 数据是: " + row);
                })
                .parse();
        if (!parse.isSuccess()) {
            // 输出验证不通过的信息
            System.out.println(parse.getMessageToString());
        }

        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }
}
```

* 解析为Java对象, 验证excel内容

```java
public class MainClass {

    @Test
    public void parseBean() {
        PoiResult<ProjectEvaluate> parse = ExcelHelper.opsParse(ProjectEvaluate.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                // 自定义列映射
                .opsColumn()
                    // 强制输入字符串, 且不能为空
                    .field("A", "projectName", "项目名称").asString().notNull().done()
                    // img类型. 导入图片必须这样写, 且字段为byte[]
                    .field("B", "img", "项目图片").asImg().done()
                    .field("C", "areaName", "所属区域").done()
                    .field("D", "province", "省份").done()
                    .field("E", "city", "市").done()
                    // 不能为空
                    .field("F", "people", "项目所属人").notNull().done()
                    // 不能为空
                    .field("G", "leader", "项目领导人").notNull().done()
                    // 必须是数字
                    .field("H", "scount", "总分").asInt().done()
                    .field("I", "avg", "历史平均分").asDouble().done()
                    .field("J", "createTime", "创建时间").asDate("yyyy-MM-dd").done()
                    .done()
                .callBack((row, index) -> {
                    // 行回调, 可以在这里改数据
                    System.out.println("当前是第:" + index + " 数据是: " + row);
                })
                .parse();
        if (!parse.isSuccess()) {
            // 输出验证不通过的信息
            System.out.println(parse.getMessageToString());
        }

        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }
}
```

* 以上如果excel数据不满足, 会收集到以下错误内容

```xml
第4行: 项目领导人-不能为空(G4)  总分-格式不正确(H4) 
第8行: 项目领导人-不能为空(G8) 
```
