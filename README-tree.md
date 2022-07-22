<details open="true">
    <summary>ExcelHelper <a>导入导出引导类</a></summary>
    <dl>
        <dd>
            <details>
                <summary>opsExport <a>导出Excel</a></summary>
                <dl>
                    <dd>
                        <details>
                            <summary>opsSheet <a>声明sheet</a></summary>
                            <dl>
                                <dd>
                                    <details>
                                        <summary>opsHeader <a>设置表头</a></summary>
                                        <dl>
                                            <dd>
                                                <details>
                                                    <summary>complex <a>复杂表头</a></summary>
                                                    <dl>
                                                        <dd>text <a>单元格声明</a></dd>
                                                    </dl>
                                                </details>
                                            </dd>
                                            <dd>
                                                <details>
                                                    <summary>simple <a>简单表头</a></summary>
                                                    <dl>
                                                        <dd>title <a>大标题</a></dd>
                                                        <dd>text <a>列标题</a></dd>
                                                        <dd>texts <a>列标题批量</a></dd>
                                                    </dl>
                                                </details>
                                            </dd>
                                            <dd>noFreeze <a>不冻结表头</a></dd>
                                        </dl>
                                    </details>
                                </dd>
                                <dd>
                                    <details>
                                        <summary>opsColumn <a>设置导出字段</a></summary>
                                        <dl>
                                            <dd>
                                                <details>
                                                    <summary>field <a>字段设置</a></summary>
                                                    <dl>
                                                        <dd>color <a>字体颜色</a></dd>
                                                        <dd>width <a>宽度</a></dd>
                                                        <dd>height <a>高度</a></dd>
                                                        <dd>wrapText <a>自动换行</a></dd>
                                                        <dd>addgn <a>水平定位</a></dd>
                                                        <dd>backColor <a>背景色</a></dd>
                                                        <dd>pattern <a>内容格式化</a></dd>
                                                        <dd>dropdown <a>下拉框</a></dd>
                                                        <dd>comment <a>注释</a></dd>
                                                        <dd>mergerRepeat <a>纵向自动合并</a></dd>
                                                        <dd>vaddgn <a>垂直定位</a></dd>
                                                        <dd>verifyIntNum <a>验证整数</a></dd>
                                                        <dd>verifyFloatNum <a>验证浮点数字</a></dd>
                                                        <dd>verifyDate <a>验证日期</a></dd>
                                                        <dd>verifyText <a>验证单元格</a></dd>
                                                        <dd>verifyCustom <a>自定义验证</a></dd>
                                                        <dd>outHandle <a>输出回调钩子</a></dd>
                                                    </dl>
                                                </details>
                                            </dd>
                                            <dd>fields <a>批量字段设置</a></dd>
                                        </dl>
                                    </details>
                                </dd>
                                <dd>
                                    <details>
                                        <summary>opsFooter <a>设置表尾</a></summary>
                                        <dl>
                                            <dd>text <a>单元格内容</a></dd>
                                        </dl>
                                    </details>
                                </dd>
                                <dd>sheetName <a>sheet名称</a></dd>
                                <dd>width <a>统一宽度</a></dd>
                                <dd>height <a>统一高度</a></dd>
                                <dd>autoNum <a>自动序号</a></dd>
                                <dd>autoNumColumnWidth <a>自动序号列宽度</a></dd>
                                <dd>mergeCells <a>批量合并单元格</a></dd>
                                <dd>mergeCellsIndex <a>批量合并单元格(下标形式)</a></dd>
                                <dd>mergeCell <a>合并单元格</a></dd>
                            </dl>
                        </details>
                    </dd>
                    <dd>parallelSheet <a>并行导出sheet</a></dd>
                    <dd>style <a>全局样式覆盖</a></dd>
                    <dd>password <a>密码设置</a></dd>
                    <dd>createBook <a>输出Workbook</a></dd>
                    <dd>fillBook <a>填充Workbook</a></dd>
                    <dd>export <a>执行导出</a></dd>
                </dl>
            </details>
        </dd>
        <dd>
            <details>
                <summary>opsReplace <a>读模板导出Excel</a></summary>
                <dl>
                    <dd>from <a>文件源</a></dd>
                    <dd>variable <a>变量替换</a></dd>
                    <dd>variables <a>批量变量替换</a></dd>
                    <dd>password <a>设置密码</a></dd>
                    <dd>replace <a>输出workbook</a></dd>
                    <dd>replaceTo <a>输出文件</a></dd>
                </dl>
            </details>
        </dd>
        <dd>
            <details>
                <summary>opsParse <a>解析Excel声明</a></summary>
                <dl>
                    <dd>from <a>文件源</a></dd>
                    <dd>
                        <details>
                            <summary>opsSheet <a>解析sheet区域声明</a></summary>
                            <dl>
                                <dd>
                                    <details>
                                        <summary>opsColumn <a>解析列定义</a></summary>
                                        <dl>
                                            <dd>
                                                <details>
                                                    <summary>field <a>字段</a></summary>
                                                    <dl>
                                                        <dd>notNdll <a>不能为空</a></dd>
                                                        <dd>asInt <a>类型int</a></dd>
                                                        <dd>asBoolean <a>类型boolean</a></dd>
                                                        <dd>asString <a>类型string</a></dd>
                                                        <dd>asLong <a>类型Long</a></dd>
                                                        <dd>asBigDecimal <a>类型Bigdecimal</a></dd>
                                                        <dd>asDate <a>类型Date</a></dd>
                                                        <dd>asDouble <a>类型Double</a></dd>
                                                        <dd>asFloat <a>类型Float</a></dd>
                                                        <dd>asImg <a>类型Img</a></dd>
                                                        <dd>asShort <a>类型Short</a></dd>
                                                        <dd>asChar <a>类型Char</a></dd>
                                                        <dd>asByCustom <a>自定义类型</a></dd>
                                                    </dl>
                                                </details>
                                            </dd>
                                        </dl>
                                    </details>
                                </dd>
                                <dd>callBack <a>解析回调钩子</a></dd>
                                <dd>parse <a>解析文件</a></dd>
                            </dl>
                        </details>
                    </dd>
                </dl>
            </details>
        </dd>
    </dl>
</details>