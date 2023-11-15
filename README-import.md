# 不验证解析
* 导入文件示例图

![image](https://user-images.githubusercontent.com/29246805/171568340-39a10aed-1c00-405c-a7f8-3157ab40f10d.png)


* 代码

```java
PoiResult<HashMap> parse = ExcelHelper.opsParse(HashMap.class)
        .from("src/test/java/excel/export/excel/simpleExport.xlsx")
        // 指定数据区域
        .opsSheet(0, 1, 0)
        .parse();
if (parse.hasError()) {
    // 输出验证不通过的信息
    System.out.println(parse.getMessageToString());
}
// 打印解析的数据
parse.getData().forEach(System.out::println);

```

# 验证解析

```java
Map<String, String> cityMapping = new HashMap<String, String>();
cityMapping.put("西安", "1");
cityMapping.put("北京", "2");

PoiResult<HashMap> parse = ExcelHelper.opsParse(HashMap.class)
    .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
    // 指定数据区域
    .opsSheet(0, 1, 0)
    // 自定义列映射
    .opsColumn()
        // 强制输入字符串, 且不能为空
        .field("A", "projectName").asString().notNull().done()
        // img类型. 导入图片必须这样写, 且字段为byte[]
        .field("B", "img").asImg().done()
        // 必须为数字
        .field("C", "areaName").asString().pattern("\\d+").done()
        .field("D", "province").asFloat().map((val)->{
            if (val > 100){
                // 会被收集异常
                throw PoiException.error("分数太大了!");
            }
            return val;
        }).done()
        // 城市，去两边空格，并映射字典。
        .field("E", "city").asString().defaultValue("未知").trim().map(cityMapping::get).done()
        // 不能为空
        .field("F", "people").asBigDecimal().notNull().done()
        // 不能为空
        .field("G", "leader").asString().notNull().done()
        // 必须是数字
        .field("H", "scount").asLong().done()
        .field("I", "avg").asByCustom((row, col, value) ->{
            if(value != null){
                // 自定义处理数据转换验证等操作
                return 1;
            }
            return value;
        }).done()
        .field("J", "createTime").asDate().pattern("yyyy-MM-dd").done()
        .done()
    // 行级别钩子
    .map((row, index) -> {
        // 行回调, 可以在这里改数据
        System.out.println("当前是第:" + index + " 数据是: " + row);
        // 也可以验证数据
        if (row.get("leader") == null){
            throw PoiException.error("leader不能为空");
        }
    })
    .parse();
if (parse.hasError()) {
    // 输出验证不通过的信息
    System.out.println(parse.getErrorInfoString());
}

// 存在解析通过的数据，打印数据
if (parse.hasData()) {
    parse.getData().forEach(System.out::println);
}
```

* 以上如果excel验证不通过, 会收集到以下错误内容

```xml
第4行: 项目领导人-不能为空(G4)  总分-格式不正确(H4) 
第8行: 项目领导人-不能为空(G8) 
```
