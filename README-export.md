选择xls还是xlsx？

> xls速度较快，单sheet最大65535行，体积大. xlsx速度慢，单sheet最大1048576行，体积小

# 简单导出

* 代码示例
> 导出图片，需要定义`data`中`img`字段为byte[]类型，自行把图片转byte[]

```java
// 指定导出XLSX格式
ExcelHelper.opsExport(PoiWorkbookType.XLSX)
        .opsSheet(data)
        .opsHeader().simple().texts("项目名称", "项目图", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间").done()
        .opsColumn().fields("projectName", "img", "areaName", "province", "city", "people", "leader", "scount", "avg", "createTime").done()
        .done()
        .export("src/test/java/excel/export/excel/simpleExport.xlsx");
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/171567064-3b62d725-96c8-4ee1-b91c-59f1d5f98f56.png)


# 简单导出 + 属性完整示例

* 代码示例

```java
// 自定义全局 样式
ICellStyle titleStyle = new ICellStyle() {

    // 样式所属：支持大标题，副标题，数据单元格，尾部数据单元格 TITLE，HEADER，CELL，FOOTER
    @Override
    public CellPosition getPosition() {
        return CellPosition.TITLE;
    }

    // 字体样式设置
    @Override
    public void handleStyle(Font font, CellStyle cellStyle) {
        font.setFontHeightInPoints((short) 20);
        // 红色字体
        font.setColor(IndexedColors.RED.index);
        // 居左
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
    }
};

ExcelHelper.opsExport(PoiWorkbookType.XLSX)
        // 全局样式覆盖, 可以传多个
        .style(titleStyle)
        // 导出添加密码
        .password("123456")
        // 并行流导出
        .parallelSheet()
        // sheet 导出声明
        .opsSheet(data)
            // 自动生成序号, （复杂表头下, 需要自己定义序号列）
            .autoNum()
            // 自定义数据行高度, 默认excel正常高度
            .height(CellPosition.CELL, 300)
            // 全局单元格宽度100000
            .width(100000)
            // 序号列宽度, 默认2000
            .autoNumColumnWidth(3000)
            // sheet名字
            .sheetName("简单导出")
            // 表头声明
            .opsHeader()
                // 不冻结表头, 默认冻结
                .noFreeze()
                // 简单模式
                .simple()
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
            // 数据列声明
            .opsColumn()
                // 批量导出字段
                .fields("projectName", "img", "areaName", "province", "people")
                // 个性化导出字段设置
                .field("city")
                // 超出宽度换行显示
                .wrapText()
                // 下拉框
                .dropdown("北京", "西安", "上海", "广州")
                // 行数据相同合并
                .mergerRepeat()
                // 行高单独设置
                .height(500)
                // 批注
                .comment("城市选择下拉框内容哦")
                // 宽度设置
                .width(6000)
                // 字段导出回调
                .map((val, row, style, rowIndex) -> {
                    // 如果是北京, 设置背景色为黄色
                    if (val.equals("北京")) {
                        style.setBackColor(IndexedColors.YELLOW);
                        style.setHeight(900);
                        style.setComment("北京搞红色");
                        // 属性值自定义
                        int index = rowIndex + 1;
                        return "=J" + index + "+K" + index;
                    }
                    return val;
                }).done()
                .field("createTime")
                // 区域相同, 合并当前列
                .mergerRepeat("areaName")
                // 格式化, 单元格格式，具体参考这个类，或者Excel表格。org.apache.poi.ss.usermodel.BuiltinFormats
                .pattern("yyyy-MM-dd")
                // 居左
                .align(HorizontalAlignment.LEFT)
                // 居中
                .valign(VerticalAlignment.CENTER)
                // 背景黄色
                .backColor(IndexedColors.YELLOW)
                // 金色字体
                .color(IndexedColors.GOLD).done()
                .fields("leader", "scount")
                .field("avg").pattern("0.00").done()
                .done()
            // 尾行设计
            .opsFooter()
                // 字符合并, 尾行合并, 行数从1开始, 会自动计算数据行
                .text("合计", "A1:H1")
                // 公式应用
                .text(String.format("=SUM(J3:J%s)", 2 + data.size()), "1,1,J,J")
                .text(String.format("=AVERAGE(K3:K%s)", 2 + data.size()), "1,1,K,K")
                // 坐标合并
                .text("作者:625", 0, 0, 8, 8).done()
                .done()
        // 执行导出
        .export("src/test/java/excel/export/excel/simpleExport2.xlsx");
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/174421969-b570cd36-a012-4035-a3f1-45a84bdd2be2.png)

# 复杂表头

* 代码示例

```java
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
                // 指定区域添加图片
                .addImage(imageParseBytes(new File("C:\\Users\\35361\\Documents\\code\\self\\poi-excel\\src\\test\\java\\excel\\export\\data\\1.jpg")), "F4:G13")
                .done()
            .export("src/test/java/excel/export/excel/complexExport.xlsx");
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/172006040-db5b4016-a54e-4816-8585-bf760a4f8e54.png)


# 复杂对象级联导出

* 代码示例

```java
ExcelHelper.opsExport(PoiWorkbookType.XLSX)
        .opsSheet(complexData)
        .opsHeader().simple().texts("學生姓名", "所在班級", "所在學校", "更多父母姓名").done()
        .opsColumn().fields("name", "classRoom.name", "classRoom.school.name", "moreInfo.parent.age").done()
        .done()
        .export("src/test/java/excel/export/excel/complexObject.xlsx");
 
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/171567981-1690ea1e-7116-40de-82b0-3ae3aa3a0754.png)

# 模板导出

* 代码示例

```java
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
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/171568101-1ab4d733-85be-4b7e-ac88-82098f790d4f.png)

# 多sheet导出

* 代码示例

```java
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
```
* 导出示例

![image](https://user-images.githubusercontent.com/29246805/171568172-4fd123d5-dc6f-49a8-965a-a7f76751fa46.png)


# 大数据导出，防OOM

* 代码示例

```java
// 声明导出BIG XLSX
ExcelHelper.opsExport(PoiWorkbookType.BIG_XLSX)
        .opsSheet(bigData)
        .sheetName("1")
        .opsHeader().simple().texts("项目名称", "项目图", "所属区域", "省份", "市", "项目所属人", "项目领导人", "得分", "平均分", "创建时间").done()
        .opsColumn().fields("projectName", "img", "areaName", "province", "city", "people", "leader", "scount", "avg", "createTime").done()
        .done()
        .export("src/test/java/excel/export/excel/bigData.xlsx");
```

# 读模板替换变量导出
![image](https://user-images.githubusercontent.com/29246805/171568236-9960f08c-782d-4457-92e8-a681cbcb06e6.png)


* 代码示例

```java
ExcelHelper.opsReplace()
        .from("src/test/java/excel/replace/replace.xlsx")
        .var("projectName","中青旅")
        .var("buildName","管材生产")
        .var("sendDate","2020-02-02")
        .var("reciveSb","张三")
        .var("phone","15594980303")
        .var("address","陕西省xxxx")
        .var("company","社保局")
        .var("remark","李四")
        .replaceTo("src/test/java/excel/replace/replace2.xlsx");
```

* 导出结果

![image](https://user-images.githubusercontent.com/29246805/171568276-76572937-b483-441c-bc53-10ea3eea0b4d.png)
