package com.github.stupdit1t.excel.core.parse;

import com.github.stupdit1t.excel.callback.InCallback;
import com.github.stupdit1t.excel.common.PoiResult;
import com.github.stupdit1t.excel.common.PoiSheetDataArea;
import com.github.stupdit1t.excel.core.AbsParent;
import com.github.stupdit1t.excel.core.ExcelUtil;
import org.apache.commons.lang3.StringUtils;

import java.util.Collections;
import java.util.Map;

/**
 * 导出规则定义
 */
public class OpsSheet<R> extends AbsParent<OpsParse<R>> {

    /**
     * 读取的sheetNum
     */
    int sheetIndex;

    /**
     * 读取的sheetName
     */
    String sheetName;

    /**
     * 选择的sheet模式
     * 1 下标
     * 2 名字
     */
    int sheetMode;

    /**
     * 表头行数量
     */
    int headerCount;

    /**
     * 尾部行数量
     */
    int footerCount;

    /**
     * 导入列定义
     */
    OpsColumn<R> opsColumn;

    /**
     * 行回调方法
     */
    InCallback<R> callback;

    /**
     * 声明构造
     *
     * @param parent 当前对象
     */
    public OpsSheet(OpsParse<R> parent, int sheetIndex, int headerCount, int footerCount) {
        super(parent);
        this.headerCount = headerCount;
        this.footerCount = footerCount;
        checkSetSheetMode(1);
        this.sheetIndex = sheetIndex;
    }

    /**
     * 声明构造
     *
     * @param parent 当前对象
     */
    public OpsSheet(OpsParse<R> parent, String sheetName, int headerCount, int footerCount) {
        super(parent);
        this.headerCount = headerCount;
        this.footerCount = footerCount;
        checkSetSheetMode(2);
        this.sheetName = sheetName;
    }

    public OpsColumn<R> opsColumn() {
        this.opsColumn = new OpsColumn<>(this);
        return this.opsColumn;
    }

    /**
     * 检测是否已经被设置状态
     */
    private void checkSetSheetMode(int wantSetMode) {
        if (sheetMode != 0 && sheetMode != wantSetMode) {
            throw new UnsupportedOperationException("仅支持设置 1 种sheet读取方式");
        }
        this.sheetMode = wantSetMode;
    }

    /**
     * 行回调方法
     *
     * @return OpsSheet
     */
    public OpsSheet<R> callBack(InCallback<R> callback) {
        this.callback = callback;
        return this;
    }

    /**
     * 解析sheet方法
     *
     * @return PoiResult
     */
    public PoiResult<R> parse() {
        Map<String, InColumn<?>> columns = Collections.emptyMap();
        if (this.opsColumn != null) {
            columns = this.opsColumn.columns;
        }

        // 校验用户输入, 必填项校验
        if (StringUtils.isBlank(this.parent.fromPath) && this.parent.fromStream == null) {
            throw new UnsupportedOperationException("Excel来源不能为空!");
        }
        // 校验用户输入, 非Map, 列必填
        if (!this.parent.mapData) {
            if (columns.isEmpty()) {
                throw new UnsupportedOperationException("导入的opsColumn字段不能为空!");
            }
        }
        PoiSheetDataArea poiSheetArea;
        if (StringUtils.isNotBlank(this.sheetName)) {
            poiSheetArea = new PoiSheetDataArea(this.sheetName, this.headerCount, this.footerCount);
        } else {
            poiSheetArea = new PoiSheetDataArea(this.sheetIndex, this.headerCount, this.footerCount);
        }
        if (this.parent.fromMode == 1) {
            if(this.parent.password != null){
                return ExcelUtil.readSheet(this.parent.fromPath, this.parent.password, poiSheetArea, columns, this.callback, this.parent.rowClass);
            }
            return ExcelUtil.readSheet(this.parent.fromPath, poiSheetArea, columns, this.callback, this.parent.rowClass);
        } else if (this.parent.fromMode == 2) {
            if(this.parent.password != null){
                return ExcelUtil.readSheet(this.parent.fromStream, this.parent.password, poiSheetArea, columns, this.callback, this.parent.rowClass);
            }
            return ExcelUtil.readSheet(this.parent.fromStream, poiSheetArea, columns, this.callback, this.parent.rowClass);
        }
        return PoiResult.fail(null);
    }

}