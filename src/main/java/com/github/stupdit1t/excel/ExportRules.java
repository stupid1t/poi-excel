package com.github.stupdit1t.excel;

import com.github.stupdit1t.excel.common.PoiCommon;
import com.github.stupdit1t.excel.style.DefaultCellStyleEnum;
import com.github.stupdit1t.excel.style.ICellStyle;

import java.util.Map;

public class ExportRules {

    /**
     * sheetName
     */

    private String sheetName;

    /**
     * 是否带序号
     */

    private boolean autoNum;

    /**
     * 列数据规则定义
     */
    private final Column[] column;

    /**
     * 表头名
     */
    private String title;

    /**
     * 标题列
     */
    private String[] header;

    /**
     * excel头：合并规则及值，rules.put("1,1,A,G", "其它应扣"); 对应excel位置
     */
    private Map<String, String> headerRules;

    /**
     * excel尾 ： 合并规则及值，rules.put("1,1,A,G", 值); 对应excel位置
     */
    private Map<String, String> footerRules;

    // --------------------无关设置字段-------------------------

    /**
     * 最大单元格列数
     */
    private int maxColumns = 0;

    /**
     * 表头最大行数
     */
    private int maxRows = 0;

    /**
     * 是否合并表头
     */
    private boolean ifMerge;

    /**
     * 是否有页脚
     */
    private boolean ifFooter;

    /**
     * 是否简单导出
     */
    private boolean simple;

    /**
     * 是否导出xlsx
     */
    private boolean xlsx = true;

    /**
     * 全局单元格样式
     */
    private ICellStyle[] globalStyle = DefaultCellStyleEnum.values();

    /**
     * 初始化规则，构建一个简单表头
     *
     * @param column 定义导出列字段
     * @param header 表头设计
     */
    public static ExportRules simpleRule(Column[] column, String[] header) {
        return new ExportRules(column, header, true);
    }

    /**
     * 初始化规则，构建一个复杂表头
     *
     * @param column      定义导出列字段
     * @param headerRules 表头设计
     */
    public static ExportRules complexRule(Column[] column, Map<String, String> headerRules) {
        return new ExportRules(column, headerRules);
    }

    /**
     * 常规一行表头构造,不带尾部
     *
     * @param column 列数据规则定义
     * @param header 表头标题
     * @param simple 简单表头
     */
    private ExportRules(Column[] column, String[] header, boolean simple) {
        super();
        this.column = column;
        this.simple = simple;
        setHeader(header);
    }

    /**
     * 复杂表头构造
     *
     * @param column      列数据规则定义
     * @param headerRules 表头设计
     */
    private ExportRules(Column[] column, Map<String, String> headerRules) {
        super();
        this.column = column;
        setHeaderRules(headerRules);
    }

    private void setHeader(String[] header) {
        this.header = header;
        this.maxRows = this.maxRows + 1;
        this.maxColumns = header.length - 1;
    }

    private void setHeaderRules(Map<String, String> headerRules) {
        this.headerRules = headerRules;
        int[] mapRowColNum = PoiCommon.getMapRowColNum(headerRules);
        this.maxRows = mapRowColNum[0];
        this.maxColumns = mapRowColNum[1];
        this.ifMerge = true;
    }

    /**
     * 尾行设计
     *
     * @param footerRules 尾部合计行设计
     */
    public ExportRules footerRules(Map<String, String> footerRules) {
        this.footerRules = footerRules;
        this.ifFooter = true;
        return this;
    }


    /**
     * sheet名
     *
     * @param sheetName sheet名字
     */
    public ExportRules sheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    /**
     * 自动生成序号
     * complexRule：需要自定义手动定义表头
     * simpleRule：自动生成表头序号
     *
     * @param autoNum 自动生成序号
     */
    public ExportRules autoNum(boolean autoNum) {
        this.autoNum = autoNum;
        if (autoNum && simple) {
            String[] headerNew = new String[this.header.length + 1];
            System.arraycopy(this.header, 0, headerNew, 1, headerNew.length - 1);
            headerNew[0] = "序号";
            this.header = headerNew;
            this.maxColumns = headerNew.length - 1;
        }

        return this;
    }

    /**
     * 表头设置
     *
     * @param title 表头标题
     */
    public ExportRules title(String title) {
        if (this.headerRules != null) {
            throw new UnsupportedOperationException("不能同时设置title和headerRules!请在headerRules设计excel标题");
        }
        this.title = title;
        this.maxRows = this.maxRows + 1;
        return this;
    }

    /**
     * 全局单元格样式设置
     *
     * @param styles 全局样式设置,传数组,不需要全部覆盖
     */
    public ExportRules globalStyle(ICellStyle... styles) {
        this.globalStyle = styles;
        return this;
    }

    public ExportRules xlsx(boolean xlsx) {
        this.xlsx = xlsx;
        return this;
    }

    boolean isXlsx() {
        return xlsx;
    }

    boolean isAutoNum() {
        return autoNum;
    }

    Column[] getColumn() {
        return column;
    }

    ICellStyle[] getGlobalStyle() {
        return globalStyle;
    }

    String getSheetName() {
        return sheetName;
    }

    int getMaxColumns() {
        return maxColumns;
    }

    int getMaxRows() {
        return maxRows;
    }

    String getTitle() {
        return title;
    }

    String[] getHeader() {
        return header;
    }

    Map<String, String> getHeaderRules() {
        return headerRules;
    }

    Map<String, String> getFooterRules() {
        return footerRules;
    }

    boolean isIfMerge() {
        return ifMerge;
    }

    boolean isIfFooter() {
        return ifFooter;
    }

}
