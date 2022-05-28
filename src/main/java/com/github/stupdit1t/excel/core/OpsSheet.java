package com.github.stupdit1t.excel.core;

import com.github.stupdit1t.excel.style.CellPosition;

import java.util.List;

/**
 * 导出规则定义
 */
public class OpsSheet<R> extends AbsParent<OpsExport> {

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
     * 尾行高度
     */
    short footerHeight = -1;

    /**
     * 是否自动带序号
     */
    boolean autoNum;

    /**
     * 自动排序列宽度
     */
    int autoNumColumnWidth = -1;

    /**
     * sheet名字
     */
    String sheetName;

    /**
     * 导出的数据
     */
    List<R> data;

    /**
     * 导出的表头定义
     */
    OpsHeader<R> opsHeader;

    /**
     * 导出的数据列定义
     */
    OpsColumn<R> opsColumn;

    /**
     * 复杂尾设计容器
     */
    OpsFooter<R> opsFooter;

    OpsSheet(OpsExport opsExport) {
        super(opsExport);
    }

    /**
     * 表头设置
     *
     * @return OpsHeader<R>
     */
    public OpsHeader<R> opsHeader() {
        this.opsHeader = new OpsHeader<>(this);
        return this.opsHeader;
    }

    /**
     * 数据列定义
     *
     * @return OpsColumn<R>
     */
    public OpsColumn<R> opsColumn() {
        this.opsColumn = new OpsColumn<>(this);
        return this.opsColumn;
    }

    /**
     * 表头设置
     *
     * @return OpsSheet<R>
     */
    public OpsFooter<R> opsFooter() {
        this.opsFooter = new OpsFooter<>(this);
        return this.opsFooter;
    }

    /**
     * sheetName 定义
     *
     * @return OpsSheet<R>
     */
    public OpsSheet<R> sheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    /**
     * sheetName 定义
     *
     * @return OpsSheet<R>
     */
    public OpsSheet<R> autoNum(boolean autoNum) {
        this.autoNum = autoNum;
        return this;
    }

    /**
     * 自动序号列宽度
     *
     * @return OpsSheet<R>
     */
    public OpsSheet<R> autoNumColumnWidth(int autoNumColumnWidth) {
        this.autoNumColumnWidth = autoNumColumnWidth;
        return this;
    }

    /**
     * sheetName 定义
     *
     * @return OpsSheet<R>
     */
    public OpsSheet<R> height(CellPosition cellPosition, int height) {
        switch (cellPosition) {
            case FOOTER:
                this.footerHeight = (short) height;
                break;
            case CELL:
                this.cellHeight = (short) height;
                break;
            case TITLE:
                this.titleHeight = (short) height;
                break;
            case HEADER:
                this.headerHeight = (short) height;
                break;
        }
        return this;
    }
}