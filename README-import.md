# 导入Map

### 自动映射列

```java
public class Test {
    public void parseMap1() {
        name.set("快速转map，自动映射列");
        PoiResult<HashMap> parse = ExcelHelper.opsParse(HashMap.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                // 自动映射列，字段值为A,B,C,D
                .opsColumn(true).done()
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }
}
```

### 自动映射列后，指定列替换

```java
public class Test {
    public void parseMap2() {
        name.set("快速转map，自动映射列，指定列替换");
        PoiResult<HashMap> parse = ExcelHelper.opsParse(HashMap.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                .opsColumn(true)
                .field(Col.H, "H列保留2位数").type(double.class).scale(2)
                .done()
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }
}
```

### 不自动映射，提取指定列

```java
public class Test {
    public void parseMap3() {
        name.set("快速转map，不自动映射，提取指定列");
        PoiResult<HashMap> parse = ExcelHelper.opsParse(HashMap.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                .opsColumn()
                // 1.去除两边空格，2.不能为空，3.默认值，4.正则校验 
                .field(Col.A, "name").trim().notNull().defaultValue("张三").regex("中青旅\\d{1}")
                // 保留2位        
                .field(Col.H, "score").scale(2)
                // 图片
                .field(Col.J, "img").type(byte[].class)
                .done()
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }
}
```

# 导入对象

### 自动映射

```java

@Test
public class Test {
    public void parseBean1() {
        name.set("自动映射列");
        PoiResult<ProjectEvaluate> parse = ExcelHelper.opsParse(ProjectEvaluate.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 1)
                .opsColumn(true).done()
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }

        // 打印解析的数据
        System.out.println("数据行数:" + parse.getData().size());
        parse.getData().forEach(System.out::println);
    }
}
```

### 自动映射转换

```java
public class Test {
    @Test
    public void parseBean2() {
        name.set("自动映射列，指定列替换");
        PoiResult<ProjectEvaluate> parse = ExcelHelper.opsParse(ProjectEvaluate.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 1)
                .opsColumn(true)
                .field(Col.H, ProjectEvaluate::getScore).type(double.class).scale(2)
                .done()
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }
}
```

### 不自动映射，提取指定列

```java
public class Test {
    @Test
    public void parseBean3() {
        name.set("不自动映射，提取指定列");
        PoiResult<ProjectEvaluate> parse = ExcelHelper.opsParse(ProjectEvaluate.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 0)
                .opsColumn()
                .field(Col.A, ProjectEvaluate::getProjectName)
                .field(Col.H, ProjectEvaluate::getScore)
                .done()
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }
}
```

### 提取指定列，校验, 转换

```java
public class Test {
    @Test
    public void parseBean4() {
        name.set("提取指定列，校验, 转换");
        Map<String, Integer> cityMapping = new HashMap<>();
        cityMapping.put("西安", 1);
        cityMapping.put("北京", 2);

        PoiResult<ProjectEvaluate> parse = ExcelHelper.opsParse(ProjectEvaluate.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 1)
                .opsColumn()
                .field(Col.A, "projectName").trim().notNull().defaultValue("张三").regex("中青旅\\d{1}")
                .field(Col.D, ProjectEvaluate::getProvince)
                // 值映射转换，也可以异常处理校验等
                .field(Col.E, "cityKey").notNull().map(cityMapping::get)
                .done()
                .parse();
        if (parse.hasError()) {
            // 输出验证不通过的信息
            System.out.println(parse.getErrorInfoString());
        }
        // 打印解析的数据
        parse.getData().forEach(System.out::println);
    }
}
```

错误信息输出如下

```shell
[A2]: 格式不正确
[A3]: 格式不正确
[A4]: 格式不正确
[A5]: 格式不正确
```

### 大数据分批处理

```java
public class Test {
    @Test
    public void parseBean5() {
        name.set("大数据分批处理");
        ExcelHelper.opsParse(ProjectEvaluate.class)
                .from("src/test/java/excel/parse/excel/simpleExport.xlsx")
                // 指定数据区域
                .opsSheet(0, 1, 1)
                .opsColumn(true).done()
                // 2个处理一次
                .parsePart(2, (result) -> {
                    if (result.hasError()) {
                        // 输出验证不通过的信息
                        System.out.println(result.getErrorInfoString());
                    }

                    // 打印解析的数据
                    System.out.println("数据行数:" + result.getData().size());
                    result.getData().forEach(System.out::println);
                });

    }
}
```
