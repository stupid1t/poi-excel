package com.github.stupdit1t.excel.core.export;

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
    private final List<OutColumn<?>> column;

    /**
     * 表头名
     */
    private String title;

    /**
     * 标题列
     */
    private LinkedHashMap<String, BiConsumer<Font, CellStyle>> simpleHeader;

    /**
     * excel头：合并规则及值，rules.put("1,1,A,G", "其它应扣"); 对应excel位置
     */
    private List<ComplexCell> complexHeader;

    /**
     * excel尾 ： 合并规则及值，rules.put("1,1,A,G", 值); 对应excel位置
     */
    private List<ComplexCell> footerRules;

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
     * Excel密码, 只支持xls 格式
     */
    private String password;

    /**
     * 是否冻结表头
     */
    private boolean freezeHeader = true;

    /**
     * 标题高度
     */
    private short titleHeight = -1;

    /**
     * 表头高度
     */
    private short headerHeight = -1;

    /**
     * 单元格高度
     */
    private short cellHeight = -1;

    /**
     * 尾部高度
     */
    private short footerHeight = -1;

    /**
     * 自动排序列宽度
     */
    private int autoNumColumnWidth = 2000;

    /**
     * 自定义合并的单元格
     */
    private List<Integer[]> mergerCells;

    /**
     * 初始化规则，构建一个简单表头
     *
     * @param column 定义导出列字段
     * @param header 表头设计
     */
    public static ExportRules simpleRule(List<OutColumn<?>> column, LinkedHashMap<String, BiConsumer<Font, CellStyle>> header) {
        return new ExportRules(column, header, true);
    }

    /**
     * 初始化规则，构建一个复杂表头
     *
     * @param column        定义导出列字段
     * @param complexHeader 表头设计
     */
    public static ExportRules complexRule(List<OutColumn<?>> column, List<ComplexCell> complexHeader) {
        return new ExportRules(column, complexHeader);
    }

    /**
     * 常规一行表头构造,不带尾部
     *
     * @param column 列数据规则定义
     * @param header 表头标题
     * @param simple 简单表头
     */
    private ExportRules(List<OutColumn<?>> column, LinkedHashMap<String, BiConsumer<Font, CellStyle>> header, boolean simple) {
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
    private ExportRules(List<OutColumn<?>> column, List<ComplexCell> complexHeader) {
        super();
        this.column = column;
        setComplexHeader(complexHeader);
    }

    /**
     * 设置简单表头
     *
     * @param simpleHeader 简单表头设置
     */
    private void setSimpleHeader(LinkedHashMap<String, BiConsumer<Font, CellStyle>> simpleHeader) {
        this.simpleHeader = simpleHeader;
        this.maxRows = this.maxRows + 1;
        this.maxColumns = simpleHeader.size() - 1;
    }

    /**
     * 设置复杂表头
     *
     * @param complexHeader 复杂表土设置
     */
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

    /**
     * 获取book密码
     *
     * @return String
     */
    public String getPassword() {
        return password;
    }

    /**
     * 是否xlsx格式
     *
     * @return boolean
     */
    public boolean isXlsx() {
        return xlsx;
    }

    /**
     * 全局样式
     *
     * @return ICellStyle[]
     */
    public ICellStyle[] getGlobalStyle() {
        return globalStyle;
    }

    /**
     * 获取sheet名字
     *
     * @return String
     */
    public String getSheetName() {
        return sheetName;
    }

    /**
     * 是否自动生成序号
     *
     * @return boolean
     */
    public boolean isAutoNum() {
        return autoNum;
    }

    /**
     * 获取输出列
     *
     * @return List<OutColumn < ?>>
     */
    public List<OutColumn<?>> getColumn() {
        return column;
    }

    /**
     * 获取大标题
     *
     * @return List<OutColumn < ?>>
     */
    public String getTitle() {
        return title;
    }

    /**
     * 获取简单表头设置
     *
     * @return LinkedHashMap<String, BiConsumer < Font, CellStyle>>
     */
    public LinkedHashMap<String, BiConsumer<Font, CellStyle>> getSimpleHeader() {
        return simpleHeader;
    }

    /**
     * 获取复杂表头设计
     *
     * @return List<ComplexCell>
     */
    public List<ComplexCell> getComplexHeader() {
        return complexHeader;
    }

    /**
     * 获取复杂尾部设计
     *
     * @return List<ComplexCell>
     */
    public List<ComplexCell> getFooterRules() {
        return footerRules;
    }

    /**
     * 获取最大列
     *
     * @return int
     */
    public int getMaxColumns() {
        return maxColumns;
    }

    /**
     * 获取最大行
     *
     * @return int
     */
    public int getMaxRows() {
        return maxRows;
    }

    /**
     * 是否合并模式
     *
     * @return boolean
     */
    public boolean isIfMerge() {
        return ifMerge;
    }

    /**
     * 是否有尾行
     *
     * @return boolean
     */
    public boolean isIfFooter() {
        return ifFooter;
    }

    /**
     * 是否简单导出
     *
     * @return boolean
     */
    public boolean isSimple() {
        return simple;
    }

    /**
     * 是否冻结表头
     *
     * @return boolean
     */
    public boolean isFreezeHeader() {
        return freezeHeader;
    }

    /**
     * 大标题高度
     *
     * @return short
     */
    public short getTitleHeight() {
        return titleHeight;
    }

    /**
     * 表头高度
     *
     * @return short
     */
    public short getHeaderHeight() {
        return headerHeight;
    }

    /**
     * 单元格高度
     *
     * @return short
     */
    public short getCellHeight() {
        return cellHeight;
    }

    /**
     * 尾行高度
     *
     * @return short
     */
    public short getFooterHeight() {
        return footerHeight;
    }

    /**
     * 自动学号列宽度
     *
     * @return int
     */
    public int getAutoNumColumnWidth() {
        return autoNumColumnWidth;
    }

    /**
     * 自定义合并单元格
     *
     * @return List<Integer[]>
     */
    public List<Integer[]> getMergerCells() {
        return mergerCells;
    }

    /**
     * 设置sheet名字
     * @param sheetName sheet名字
     */
    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    /**
     * 大标题设置
     *
     * @param title 标题
     */
    public void setTitle(String title) {
        this.title = title;
    }

    /**
     * 是否导出07 xlsx
     *
     * @param xlsx true or false
     */
    public void setXlsx(boolean xlsx) {
        this.xlsx = xlsx;
    }

    /**
     * 设置全局样式
     *
     * @param globalStyle 样式
     */
    public void setGlobalStyle(ICellStyle... globalStyle) {
        this.globalStyle = globalStyle;
    }

    /**
     * 设置excel密码
     *
     * @param password 密码
     */
    public void setPassword(String password) {
        this.password = password;
    }

    /**
     * 是否冻结表头
     *
     * @param freezeHeader 冻结
     */
    public void setFreezeHeader(boolean freezeHeader) {
        this.freezeHeader = freezeHeader;
    }

    /**
     * 大标题高度
     *
     * @param titleHeight 高度
     */
    public void setTitleHeight(short titleHeight) {
        this.titleHeight = titleHeight;
    }

    /**
     * 表头高度
     *
     * @param headerHeight 高度
     */
    public void setHeaderHeight(short headerHeight) {
        this.headerHeight = headerHeight;
    }

    /**
     * 单元格高度
     *
     * @param cellHeight 高度
     */
    public void setCellHeight(short cellHeight) {
        this.cellHeight = cellHeight;
    }

    /**
     * 尾行高度
     *
     * @param footerHeight 高度
     */
    public void setFooterHeight(short footerHeight) {
        this.footerHeight = footerHeight;
    }

    /**
     * 序号列宽度
     *
     * @param autoNumColumnWidth 宽度
     */
    public void setAutoNumColumnWidth(int autoNumColumnWidth) {
        this.autoNumColumnWidth = autoNumColumnWidth;
    }

    /**
     * 设置合并单元格
     *
     * @param mergerCells 合并单元格
     */
    public void setMergerCells(List<Integer[]> mergerCells) {
        this.mergerCells = mergerCells;
    }
}
