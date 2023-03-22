package com.github.stupdit1t.excel.callback;


import com.github.stupdit1t.excel.core.export.OutColumn;

/**
 * 导出回调函数
 *
 * @author 625
 */

@FunctionalInterface
public interface OutCallback<R> {
    /**
     * 导出回调
     *
     * @param value 当前单元格值
     * @param row   当前行记录
     * @param style 自定义单元格样式
     * @param rowIndex 数据下标
     * @return 返回重置后的单元格值
     */
    Object callback(Object value, R row, OutColumn.Style style, int rowIndex);
}
