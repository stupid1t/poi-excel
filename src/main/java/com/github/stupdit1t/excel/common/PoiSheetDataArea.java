package com.github.stupdit1t.excel.common;

/**
 * 导入 sheet数据区域
 */
public class PoiSheetDataArea {

    /**
     * sheet 下标
     */
    private int sheetIndex;

    /**
     * sheet名字
     */
    private String sheetName;

    /**
     * 头部非数据行数量
     */
    private final int headerRowCount;

    /**
     * 尾部非数据行数量
     */
    private final int footerRowCount;

    /**
     * sheet数据区域
     *
     * @param sheetIndex     sheet 下标
     * @param headerRowCount 头部非数据行数量
     * @param footerRowCount 尾部非数据行数量
     */
    public PoiSheetDataArea(int sheetIndex, int headerRowCount, int footerRowCount) {
        this.sheetIndex = sheetIndex;
        this.headerRowCount = headerRowCount;
        this.footerRowCount = footerRowCount;
    }

    /**
     * sheet数据区域
     *
     * @param sheetName      sheet 名字
     * @param headerRowCount 头部非数据行数量
     * @param footerRowCount 尾部非数据行数量
     */
    public PoiSheetDataArea(String sheetName, int headerRowCount, int footerRowCount) {
        this.sheetName = sheetName;
        this.headerRowCount = headerRowCount;
        this.footerRowCount = footerRowCount;
    }

    /**
     * 获取sheet下标
     *
     * @return int
     */
    public int getSheetIndex() {
        return sheetIndex;
    }

    /**
     * 获取sheet名字
     *
     * @return int
     */
    public String getSheetName() {
        return sheetName;
    }

    /**
     * 头部非数据行数量
     *
     * @return int
     */
    public int getHeaderRowCount() {
        return headerRowCount;
    }

    /**
     * 尾部非数据行数量
     *
     * @return int
     */
    public int getFooterRowCount() {
        return footerRowCount;
    }
}
