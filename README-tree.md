<details open="true">
    <summary>ExcelHelper 结构概览 <em>点击展开</em></summary>
    <dl>
        <dd>
            <details>
                <summary>opsExport <em>导出Excel</em></summary>
                <dl>
                    <dd>
                        <details>
                            <summary>opsSheet <em>声明sheet</em></summary>
                            <dl>
                                <dd>
                                    <details>
                                        <summary>opsHeader <em>设置表头</em></summary>
                                        <dl>
                                            <dd>
                                                <details>
                                                    <summary>complex <em>复杂表头</em></summary>
                                                    <dl>
                                                        <dd>text <em>单元格声明</em></dd>
                                                    </dl>
                                                </details>
                                            </dd>
                                            <dd>
                                                <details>
                                                    <summary>simple <em>简单表头</em></summary>
                                                    <dl>
                                                        <dd>title <em>大标题</em></dd>
                                                        <dd>text <em>列标题</em></dd>
                                                        <dd>texts <em>列标题批量</em></dd>
                                                    </dl>
                                                </details>
                                            </dd>
                                            <dd>noFreeze <em>不冻结表头</em></dd>
                                        </dl>
                                    </details>
                                </dd>
                                <dd>
                                    <details>
                                        <summary>opsColumn <em>设置导出字段</em></summary>
                                        <dl>
                                            <dd>
                                                <details>
                                                    <summary>field <em>字段设置</em></summary>
                                                    <dl>
                                                        <dd>color <em>字体颜色</em></dd>
                                                        <dd>width <em>宽度</em></dd>
                                                        <dd>height <em>高度</em></dd>
                                                        <dd>wrapText <em>自动换行</em></dd>
                                                        <dd>addgn <em>水平定位</em></dd>
                                                        <dd>backColor <em>背景色</em></dd>
                                                        <dd>pattern <em>内容格式化</em></dd>
                                                        <dd>dropdown <em>下拉框</em></dd>
                                                        <dd>comment <em>注释</em></dd>
                                                        <dd>mergerRepeat <em>纵向自动合并</em></dd>
                                                        <dd>vaddgn <em>垂直定位</em></dd>
                                                        <dd>verifyIntNum <em>验证整数</em></dd>
                                                        <dd>verifyFloatNum <em>验证浮点数字</em></dd>
                                                        <dd>verifyDate <em>验证日期</em></dd>
                                                        <dd>verifyText <em>验证单元格</em></dd>
                                                        <dd>verifyCustom <em>自定义验证</em></dd>
                                                        <dd>outHandle <em>输出回调钩子</em></dd>
                                                    </dl>
                                                </details>
                                            </dd>
                                            <dd>fields <em>批量字段设置</em></dd>
                                        </dl>
                                    </details>
                                </dd>
                                <dd>
                                    <details>
                                        <summary>opsFooter <em>设置表尾</em></summary>
                                        <dl>
                                            <dd>text <em>单元格内容</em></dd>
                                        </dl>
                                    </details>
                                </dd>
                                <dd>sheetName <em>sheet名称</em></dd>
                                <dd>width <em>统一宽度</em></dd>
                                <dd>height <em>统一高度</em></dd>
                                <dd>autoNum <em>自动序号</em></dd>
                                <dd>autoNumColumnWidth <em>自动序号列宽度</em></dd>
                                <dd>mergeCells <em>批量合并单元格</em></dd>
                                <dd>mergeCellsIndex <em>批量合并单元格(下标形式)</em></dd>
                                <dd>mergeCell <em>合并单元格</em></dd>
                            </dl>
                        </details>
                    </dd>
                    <dd>parallelSheet <em>并行导出sheet</em></dd>
                    <dd>style <em>全局样式覆盖</em></dd>
                    <dd>password <em>密码设置</em></dd>
                    <dd>createBook <em>输出Workbook</em></dd>
                    <dd>fillBook <em>填充Workbook</em></dd>
                    <dd>export <em>执行导出</em></dd>
                </dl>
            </details>
        </dd>
        <dd>
            <details>
                <summary>opsReplace <em>读模板导出Excel</em></summary>
                <dl>
                    <dd>from <em>文件源</em></dd>
                    <dd>variable <em>变量替换</em></dd>
                    <dd>variables <em>批量变量替换</em></dd>
                    <dd>password <em>设置密码</em></dd>
                    <dd>replace <em>输出workbook</em></dd>
                    <dd>replaceTo <em>输出文件</em></dd>
                </dl>
            </details>
        </dd>
        <dd>
            <details>
                <summary>opsParse <em>解析Excel声明</em></summary>
                <dl>
                    <dd>from <em>文件源</em></dd>
                    <dd>
                        <details>
                            <summary>opsSheet <em>解析sheet区域声明</em></summary>
                            <dl>
                                <dd>
                                    <details>
                                        <summary>opsColumn <em>解析列定义</em></summary>
                                        <dl>
                                            <dd>
                                                <details>
                                                    <summary>field <em>字段</em></summary>
                                                    <dl>
                                                        <dd>notNdll <em>不能为空</em></dd>
                                                        <dd>asInt <em>类型int</em></dd>
                                                        <dd>asBoolean <em>类型boolean</em></dd>
                                                        <dd>asString <em>类型string</em></dd>
                                                        <dd>asLong <em>类型Long</em></dd>
                                                        <dd>asBigDecimal <em>类型Bigdecimal</em></dd>
                                                        <dd>asDate <em>类型Date</em></dd>
                                                        <dd>asDouble <em>类型Double</em></dd>
                                                        <dd>asFloat <em>类型Float</em></dd>
                                                        <dd>asImg <em>类型Img</em></dd>
                                                        <dd>asShort <em>类型Short</em></dd>
                                                        <dd>asChar <em>类型Char</em></dd>
                                                        <dd>asByCustom <em>自定义类型</em></dd>
                                                    </dl>
                                                </details>
                                            </dd>
                                        </dl>
                                    </details>
                                </dd>
                                <dd>callBack <em>解析回调钩子</em></dd>
                                <dd>parse <em>解析文件</em></dd>
                            </dl>
                        </details>
                    </dd>
                </dl>
            </details>
        </dd>
    </dl>
</details>