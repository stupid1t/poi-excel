![logo](https://user-images.githubusercontent.com/29246805/180177278-d83c1d09-5714-4f4f-8e88-4202a584b872.png)

[![OSCS Status](https://www.oscs1024.com/platform/badge/stupdit1t/poi-excel.svg?size=small)](https://www.oscs1024.com/project/stupdit1t/poi-excel?ref=badge_small)
<img alt="GitHub code size in bytes" src="https://img.shields.io/github/languages/code-size/stupdit1t/poi-excel">
<a target="_blank" href="LICENSE"><img src="https://img.shields.io/:license-MIT-blue.svg"></a>
<a target="_blank" href="https://www.oracle.com/technetwork/java/javase/downloads/index.html"><img src="https://img.shields.io/badge/JDK-1.8+-green.svg" /></a>
<a target="_blank" href="https://poi.apache.org/download.html"><img src="https://img.shields.io/badge/POI-5.2.2+-green.svg" /></a>
<a target="_blank" href='https://github.com/stupdit1t/poi-excel'><img src="https://img.shields.io/github/stars/stupdit1t/poi-excel.svg?style=social"/>
<a href='https://gitee.com/stupid1t/poi-excel/stargazers'><img src='https://gitee.com/stupid1t/poi-excel/badge/star.svg?theme=white' alt='star'></img></a>

## 一. 快速开始

**已上传maven中央仓库, 无需下载此项目, 快照版本每个月初发release**

### POM中maven直接引入

```xml
<!-- excel导入导出 POI版本为5.2.2 -->
<dependency>
    <groupId>com.github.stupdit1t</groupId>
    <artifactId>poi-excel</artifactId>
    <version>3.2.0-SNAPSHOT</version>
</dependency>
```

### 兼容两个低版本POI(截至 3.1.5)

```xml
<!-- excel导入导出 POI版本为3.17 -->
<dependency>
    <groupId>com.github.stupdit1t</groupId>
    <artifactId>poi-excel</artifactId>
    <version>poi-317.5</version>
</dependency>

<!-- excel导入导出 POI版本为4.1.2 -->
<dependency>
    <groupId>com.github.stupdit1t</groupId>
    <artifactId>poi-excel</artifactId>
    <version>poi-412.5</version>
</dependency>
```

### Spring使用示例

```java
@ApiOperation("导出异常日志")
@GetMapping("/export")
public void export(HttpServletResponse response,SysErrorLogQueryParam queryParams){
        // 1.获取列表数据
        List<SysErrorLog> data=sysErrorLogService.selectListPC(queryParams);

        // 2.执行导出
        ExcelHelper.opsExport(PoiWorkbookType.XLSX)
        .opsSheet(data)
        .opsHeader().simple()
        .texts("请求地址","请求方式","IP地址","简要信息","异常时间","创建人").done()
        .opsColumn()
        .fields("requestUri","requestMethod","ip","errorSimpleInfo","createDate","creatorName").done()
        .done()
        .export(response,"异常日志.xlsx");
        }
```

## 二. 项目优势

- 简单快速上手，且满足绝大多数业务场景
- 屏蔽POI细节，学习成本低。
- 未使用注解方式实现，纯编码代码块，去除烦人的各种POJO
- 功能强大，导入支持严格的单元格校验，导出支持公式，复杂表头和尾部设计，以及单元格样式自定义支持
- 支持读取复杂Excel模板,替换变量输出Excel,变量用${}替代

## 三. 更新记录

> 有需求才有进步，这个轮子本身就是从0开始因为需求慢慢叠加起来的。有新需求提出来,我觉得合适会更新. 如有疑问可加群帮解答: 811606008

### v3.2.0 (相对3.1.x，有部分不兼容改动)

1. 重构解析表格异常收集，a.提供行级别异常输出，b.单元格级别异常输出，c.自定义异常输入
2. 新增解析数字格式的单元格，用String接收带小数点.0的问题，提供intStr()参数，转换为整形

### v3.1.5

1. xls格式导出下拉框不能支持太多数据，更换为引用支持更多的数据
2. 列数太多，引用单元格BUG处理

### v3.1.4

1. 导出导入回调注释添加
2. 导入field方法支持只传入列和字段，不需要title
3. 导出SXSSFWorkbook格式删除临时文件

### v3.1.3

1. 解析Excel遇到未知异常捕获至PoiResult
2. 解析Excel链式方法调整，新增defaultValue
3. 增加大数据事件流分批导入功能，如一次处理百万数据，避免直接读取到内存OOM问题

### v3.1.2

1. 解析回调处理步骤POI Exception
2. 添加支持非1904日期的识别
3. 解析列添加trim方法

### v3.1.1

1. 导出支持读取Excel追加sheet页

### v3.1.0

1. 支持单元格级别的批注功能, 参考simpleExport2

### v3.0.9

1. 表头相同名字重复设置报错, fixbug

### v3.0.8

1. 导出添加设置列换行显示属性 参考简单导出simpleExport2
2. 添加sheet设置全局的单元格宽度属性 参考简单导出simpleExport2

### v3.0.7

1. POI版本升级 5.1.0 ----- 5.2.2

### v3.0.6

1. 增加导出自动感知行数据合并行功能, 方法为mergerRepeat, 参考 简单导出 + 自定义属性完整示例

### v3.0.5

1. 导出参数为空检验
2. 部分方法名调整

### v3.0.4

1. 日期格式化方法名修改 UPDATE
2. 保留导出workbook, 提供灵活性 NEW
3. 支持xlsx添加密码 NEW

### v3.0.2 ( 不兼容[1.x.x](README-1.x.md) 和 [2.x.x](README-2.x.md) 版本)

1. 提供ExcelHepler链式构建类, 帮助快捷构建. 本身还是调用ExcelUtil类
2. 优化代码结构和层次
3. 提供更精确的单元格样式控制

## 四. 导出

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

        // 指定导出XLS格式,
        ExcelHelper.opsExport(PoiWorkbookType.XLS)
                // 全局样式覆盖
                .style(titleStyle)
                // 导出添加密码
                .password("123456")
                // sheet声明
                .opsSheet(data)
                // 自动生成序号, 此功能在复杂表头下, 需要自己定义序号列
                .autoNum()
                // 自定义数据行高度, 默认excel正常高度
                .height(CellPosition.CELL, 300)
                // 全局单元格宽度100000
                .width(10000)
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
                // 当前行数据相同合并
                .mergerRepeat()
                // 超出宽度换行显示
                .wrapText()
                // 下拉框
                .dropdown("北京", "西安", "上海", "广州")
                // 行高单独设置
                .height(500)
                // 批注
                .comment("城市选择下拉框内容哦")
                // 宽度设置
                .width(6000)
                // 字段导出回调
                .outHandle((val, row, style, index) -> {
                    // 如果是北京, 设置背景色为黄色
                    if (val.equals("北京")) {
                        style.setBackColor(IndexedColors.YELLOW);
                        style.setHeight(900);
                        style.setComment("自定义设置样式批注");
                        // 属性值自定义
                        return val + "(自定义)";
                    }
                    return val;
                })
                .done()
                .field("createTime")
                // 区域相同, 合并时间
                .mergerRepeat("areaName")
                // 日期格式化
                .pattern("yyyy-MM-dd")
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

![image](https://user-images.githubusercontent.com/29246805/174421969-b570cd36-a012-4035-a3f1-45a84bdd2be2.png)

##### 3. 复杂表头导出

* 代码示例

```java
class a {

    @Test
    public void complexExport() {
        ExcelHelper.opsExport(PoiWorkbookType.XLSX)
                .opsSheet(data)
                .autoNum()
                .opsHeader()
                // 不冻结表头, 默认冻结
                .noFreeze()
                // 复杂表头模式, 支持三种合并方式, 1数字坐标 2字母坐标 3Excel坐标
                .complex()
                // excel坐标
                .text("项目资源统计", "A1:K1")
                // 字母坐标
                .text("序号", "2,3,A,A")
                // 数字坐标
                .text("基本信息", 1, 1, 1, 4)
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
                .field("宝宝生日").pattern("yyyy-MM-dd").verifyDate("2000-01-01~3000-12-31").done()
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
                // 多线程导出多sheet, 默认单线程
                .parallelSheet()
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
##### 5. 支持数据处理，如设置默认值/转换/去空格/日期格式化/excel日期识别/正则验证/自定义转换验证
##### 6. 支持大数据导入，数据分批处理，防OOM

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
        name.set("parseMap1");
        PoiResult<HashMap> parse = ExcelHelper.opsParse(HashMap.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        if (parse.hasData()) {
            result.getData().forEach(System.out::println);
        }
    }
}
```

* 解析大数据，分批处理，

```java
public class MainClass {
    @Test
    public void parseMapBig() {
        name.set("parseMapBig");
        ExcelHelper.opsParse(HashMap.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 1)
                // 每1000处理一次
                .parsePart(1000, (result) -> {
                    if (result.hasError()) {
                        // 输出验证不通过的信息
                        System.out.println(result.getErrorInfoString());
                    }
                    // 打印解析的数据
                    if (result.hasData()) {
                        result.getData().forEach(System.out::println);
                    }
                });
    }
}
```

* 解析为Java对象, 验证excel内容

```java
public class MainClass {

    @Test
    public void parseBean() {
        name.set("parseBean");
        PoiResult<ProjectEvaluate> result = ExcelHelper.opsParse(ProjectEvaluate.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                // 自定义列映射
                .opsColumn()
                // 强制输入字符串, 且不能为空
                .field("A", "projectName").asByCustom((row, col, val) -> {
                    if ("中青旅0".equals(val)) {
                        throw new PoiException("数据有误!");
                    }
                    if ("中青旅1".equals(val)) {
                        System.out.println(1 / 0);
                    }
                    // 重写值
                    return val + "-Hello";
                }).notNull().done()
                // img类型. 导入图片必须这样写, 且字段为byte[]
                .field("B", "img").asImg().done()
                .field("C", "areaName").asString().done()
                .field("D", "province").done()
                .field("E", "city").done()
                // 不能为空
                .field("F", "people").asString().pattern("\\d+").defaultValue("1").notNull().done()
                // 不能为空
                .field("G", "leader").asString().defaultValue("巨无霸").done()
                // 字符串整形数字，Excel数子类型会带浮点
                .field("H", "scount").asString().intStr().done()
                .field("I", "avg").asDouble().notNull().done()
                .field("J", "createTime").asDate().pattern("yyyy/MM/dd").trim().done()
                .done()
                .callBack((row, index) -> {
                    // 行回调, 可以在这里改数据
                    System.out.println("当前是第:" + index + " 数据是: " + row);
                    if ("中青旅2-Hello".equals(row.getProjectName())) {
                        throw new NullPointerException();
                    }
                })
                .parse();
    }
}
```

* 以上如果excel验证不满足, 收集错误内容

```xml
if(result.hasError()){
    System.out.println("===============单元格错误=================");
    String errorInfoString = result.getErrorInfoString();
    System.out.println(errorInfoString);

    System.out.println("===============数据行错误=================");
    String errorInfoLineString = result.getErrorInfoLineString();
    System.out.println(errorInfoLineString);

    // 获取原始的异常信息
    List<ErrorMessage> error = result.getError();
    // 获取原始的单元格错误
    List<String> errorInfo = result.getErrorInfo();
    // 获取原始的数据行错误
    List<String> errorInfoLine = result.getErrorInfoLine();
}
```

* 输出
```xml
===============单元格错误=================
[A2]: 数据有误!
[F2]: 格式不正确
[A3]: / by zero
[F3]: 格式不正确
[F4]: 格式不正确
[第4行]: null
[F5]: 格式不正确
[J5]: Unable to parse the date: 2022/6/53的
[F6]: 格式不正确
[F7]: 格式不正确
[F8]: 格式不正确
[F9]: 格式不正确
[F10]: 格式不正确
[F11]: 格式不正确
[F12]: 不能为空
[I12]: 不能为空
===============数据行错误=================
[第2行]: A2-数据有误! F2-格式不正确
[第3行]: A3-/ by zero F3-格式不正确
[第4行]: F4-格式不正确 null
[第5行]: F5-格式不正确 J5-Unable to parse the date: 2022/6/53的
[第6行]: F6-格式不正确
[第7行]: F7-格式不正确
[第8行]: F8-格式不正确
[第9行]: F9-格式不正确
[第10行]: F10-格式不正确
[第11行]: F11-格式不正确
[第12行]: F12-不能为空 I12-不能为空
[ parseBean ] 耗时: 934

Process finished with exit code 0

```
