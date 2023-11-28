# 导入Map

### 自动映射列
```java
 public void parseMap1(){
        name.set("快速转map，自动映射列");
        PoiResult<HashMap> parse=ExcelHelper.opsParse(HashMap.class)
        .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
        // 指定数据区域
        .opsSheet(0,1,0)
        // 自动映射列，字段值为A,B,C,D
        .opsColumn(true).done()
        .parse();
        if(parse.hasError()){
        // 输出验证不通过的信息
        System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
    parse.getData().forEach(System.out::println);
}
```

### 自动映射列后，指定列替换

```java
public void parseMap2(){
        name.set("快速转map，自动映射列，指定列替换");
        PoiResult<HashMap> parse=ExcelHelper.opsParse(HashMap.class)
        .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
        // 指定数据区域
        .opsSheet(0,1,0)
        .opsColumn(true)
        .field(Col.H,"H列保留2位数").type(double.class).scale(2)
        .done()
        .parse();
        if(parse.hasError()){
        // 输出验证不通过的信息
        System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
        }
```

### 不自动映射，提取指定列

```java
 public void parseMap3(){
        name.set("快速转map，不自动映射，提取指定列");
        PoiResult<HashMap> parse=ExcelHelper.opsParse(HashMap.class)
        .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
        // 指定数据区域
        .opsSheet(0,1,0)
        .opsColumn()
        // 1.去除两边空格，2.不能为空，3.默认值，4.正则校验 
        .field(Col.A,"name").trim().notNull().defaultValue("张三").regex("中青旅\\d{1}")
        // 保留2位        
        .field(Col.H,"score").scale(2)
        // 图片
        .field(Col.J,"img").type(byte[].class)
        .done()
        .parse();
        if(parse.hasError()){
        // 输出验证不通过的信息
        System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
        }
```

