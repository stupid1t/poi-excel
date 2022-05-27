package com.github.stupdit1t.excel.core;

import com.github.stupdit1t.excel.common.PoiCommon;
import com.github.stupdit1t.excel.style.DefaultCellStyleEnum;
import com.github.stupdit1t.excel.style.ICellStyle;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.stream.Collectors;

class ExportRules {

    /**
     * sheetName
     */

    String sheetName;

    /**
     * 是否带序号
     */

    boolean autoNum;

    /**
     * 列数据规则定义
     */
    final List<Column<?>> column;

    /**
     * 表头名
     */
    String title;

    /**
     * 标题列
     */
    LinkedHashMap<String, BiConsumer<Font, CellStyle>> simpleHeader;

    /**
     * excel头：合并规则及值，rules.put("1,1,A,G", "其它应扣"); 对应excel位置
     */
    List<ComplexCell> complexHeader;

    /**
     * excel尾 ： 合并规则及值，rules.put("1,1,A,G", 值); 对应excel位置
     */
    List<ComplexCell> footerRules;

    // --------------------无关设置字段-------------------------

    /**
     * 最大单元格列数
     */
    int maxColumns = 0;

    /**
     * 表头最大行数
     */
    int maxRows = 0;

    /**
     * 是否合并表头
     */
    boolean ifMerge;

    /**
     * 是否有页脚
     */
    boolean ifFooter;

    /**
     * 是否简单导出
     */
    boolean simple;

    /**
     * 是否导出xlsx
     */
    boolean xlsx = true;

    /**
     * 全局单元格样式
     */
    ICellStyle[] globalStyle = DefaultCellStyleEnum.values();

    /**
     * Excel密码, 只支持xls 格式
     */
    String password;

    /**
     * 是否冻结表头
     */
    boolean freezeHeader = true;

    /**
     * 标题高度
     */
    short titleHeight = -1;

    /**
     * 表头高度
     */
    short headerHeight = -1;

    /**
     * 单元格高度
     */
    short cellHeight = -1;

    /**
     * 尾部高度
     */
    short footerHeight = -1;

    /**
     * 自动排序列宽度
     */
    int autoNumColumnWidth = -1;

    /**
     * 初始化规则，构建一个简单表头
     *
     * @param column 定义导出列字段
     * @param header 表头设计
     */
    public static ExportRules simpleRule(List<Column<?>> column, LinkedHashMap<String, BiConsumer<Font, CellStyle>> header) {
        return new ExportRules(column, header, true);
    }

    /**
     * 初始化规则，构建一个复杂表头
     *
     * @param column        定义导出列字段
     * @param complexHeader 表头设计
     */
    public static ExportRules complexRule(List<Column<?>> column, List<ComplexCell> complexHeader) {
        return new ExportRules(column, complexHeader);
    }

    /**
     * 常规一行表头构造,不带尾部
     *
     * @param column 列数据规则定义
     * @param header 表头标题
     * @param simple 简单表头
     */
    private ExportRules(List<Column<?>> column, LinkedHashMap<String, BiConsumer<Font, CellStyle>> header, boolean simple) {
        super();
        this.column = column;
        this.simple = simple;
        setSimpleHeader(header);
    }

    /**
     * 复杂表头构造
     *
     * @param column        列数据规则定义
     * @param complexHeader 表头设计
     */
    private ExportRules(List<Column<?>> column, List<ComplexCell> complexHeader) {
        super();
        this.column = column;
        setComplexHeader(complexHeader);
    }

    private void setSimpleHeader(LinkedHashMap<String, BiConsumer<Font, CellStyle>> simpleHeader) {
        this.simpleHeader = simpleHeader;
        this.maxRows = this.maxRows + 1;
        this.maxColumns = simpleHeader.size() - 1;
    }

    private void setComplexHeader(List<ComplexCell> complexHeader) {
        this.complexHeader = complexHeader;
        List<Integer[]> indexLocation = complexHeader.stream().map(ComplexCell::getLocationIndex).collect(Collectors.toList());
        int[] mapRowColNum = PoiCommon.getMapRowColNum(indexLocation);
        this.maxRows = mapRowColNum[0];
        this.maxColumns = mapRowColNum[1];
        this.ifMerge = true;
    }

    /**
     * 尾行设计
     *
     * @param footerRules 尾部合计行设计
     */
    public void setFooterRules(List<ComplexCell> footerRules) {
        this.footerRules = footerRules;
        this.ifFooter = true;
    }

    /**
     * 自动生成序号
     * complexRule：需要自定义手动定义表头
     * simpleRule：自动生成表头序号
     *
     * @param autoNum 自动生成序号
     */
    public void setAutoNum(boolean autoNum) {
        this.autoNum = autoNum;
        if (autoNum && simple) {
            LinkedHashMap<String, BiConsumer<Font, CellStyle>> newHeader = new LinkedHashMap<>(this.simpleHeader.size() + 1);
            newHeader.put("序号", null);
            newHeader.putAll(this.simpleHeader);
            this.maxColumns = newHeader.size() - 1;
            this.simpleHeader = newHeader;
        }
    }

    /**
     * 表头设置
     *
     * @param title 表头标题
     */
    public void title(String title) {
        if (this.complexHeader != null) {
            throw new UnsupportedOperationException("不能同时设置title和headerRules!请在headerRules设计excel标题");
        }
        if (StringUtils.isBlank(title)) {
            return;
        }
        this.title = title;
        this.maxRows = this.maxRows + 1;
    }
}
