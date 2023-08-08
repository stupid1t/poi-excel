[![OSCS Status](https://www.oscs1024.com/platform/badge/stupdit1t/poi-excel.svg?size=small)](https://www.oscs1024.com/project/stupdit1t/poi-excel?ref=badge_small)
<img alt="GitHub code size in bytes" src="https://img.shields.io/github/languages/code-size/stupdit1t/poi-excel">
<a target="_blank" href="LICENSE"><img src="https://img.shields.io/:license-MIT-blue.svg"></a>
<a target="_blank" href="https://www.oracle.com/technetwork/java/javase/downloads/index.html"><img src="https://img.shields.io/badge/JDK-1.8+-green.svg" /></a>
<a target="_blank" href="https://poi.apache.org/download.html"><img src="https://img.shields.io/badge/POI-5.2.2+-green.svg" /></a>
<a target="_blank" href='https://github.com/stupdit1t/poi-excel'><img src="https://img.shields.io/github/stars/stupdit1t/poi-excel.svg?style=social"/>
<a href='https://gitee.com/stupid1t/poi-excel/stargazers'><img src='https://gitee.com/stupid1t/poi-excel/badge/star.svg?theme=white' alt='star'></img></a>
# poi-excel

poi-excel 是一个基于 Apache POI 的 Java 工具，旨在简化新手在处理 Excel 表格时的操作。它提供了简单、快速上手的方式，使新手能够轻松处理复杂的表格。

## 解决的问题

许多新手在使用 Apache POI 时会面临寻找正确的 API 和编写大量代码的难题。poi-excel 旨在解决这些问题，让新手可以简单轻松地完成复杂的表格处理。

## 主要特性

- **纯编码实现**：采用纯编码实现，无需使用注解，无侵入代码。这使得编写逻辑代码更加方便，同时提供了更好的复用性。

- **导入功能强大**：支持单元格级别的校验和错误输出。它能够处理大数据批处理，支持数据转换、默认值设置、图片等功能，满足各种导入需求。

- **导出功能全面**：提供了强大的导出功能。您可以轻松设计傻瓜式的表头，自定义单元格样式，公式，添加合计行、序号、图片等元素，满足各种导出需求。

- **读模板替换变量**：提供了简单的读模板功能，您可以通过替换字符和图片的方式，灵活地替换 Excel 模板中的变量。

## 最佳实践
> 需要 Java 8 环境。

只需要将以下依赖项添加到项目的 pom.xml 文件中即可：
```xml
<!-- excel导入导出 POI版本为5.2.2 -->
<dependency>
    <groupId>com.github.stupdit1t</groupId>
    <artifactId>poi-excel</artifactId>
    <version>3.2.2</version>
</dependency>
```
如版本冲突，目前兼容两个低版本POI
```xml
<!-- excel导入导出 POI版本为3.17 -->
<dependency>
<groupId>com.github.stupdit1t</groupId>
<artifactId>poi-excel</artifactId>
<version>poi-317.8</version>
</dependency>

<!-- excel导入导出 POI版本为4.1.2 -->
<dependency>
<groupId>com.github.stupdit1t</groupId>
<artifactId>poi-excel</artifactId>
<version>poi-412.8</version>
</dependency>
```

在 Spring 环境下的以下是一个简单的示例代码，进行导出操作：
```java
@GetMapping("/export")
public void export(HttpServletResponse response, SysErrorLogQueryParam queryParams) {
    // 1.获取列表数据
    List<SysErrorLog> data = ....
    
    // 2.执行导出
    ExcelHelper.opsExport(PoiWorkbookType.XLSX)
            .opsSheet(data)
            .opsHeader().simple()
                .texts("请求地址", "请求方式", "IP地址", "简要信息", "异常时间", "创建人").done()
            .opsColumn()
                .fields("requestUri","requestMethod","ip","errorSimpleInfo","createDate","creatorName").done()
            .done()
            .export(response, "异常日志.xlsx");
}
```

## 详细使用方法

请参考以下示例代码来了解如何使用poi-excel工具：

- [导出最佳实践](./README-export.md)
- [导入最佳实践](./README-import.md)

## 更新记录
[详见README-history.md](./README-history.md)

## 报告问题和寻求支持

如果您在使用 poi-excel 过程中遇到任何问题或有任何想法和建议，可以直接提出ISSUE，或您可以加入QQ 群一起探讨。QQ群号：811606008。

## 开放协议

poi-excel 使用 MIT License 开放协议，您可以自由使用、修改和分发该工具，详细的协议内容请查阅项目中的 LICENSE 文件。

让 poi-excel 成为您处理 Excel 表格的首选工具，让您的 Excel 处理任务变得简单高效！

感谢您对 poi-excel 的支持和使用！