package com.github.stupdit1t.excel.callback;


/**
 * 导出回调函数
 *
 * @author 625
 */

@FunctionalInterface
public interface CellCallback {

    /**
     * 导出回调
     *
     * @param value 当前单元格值
     * @param row   行号
     * @param col   列号
     * @return 返回重置后的单元格值
     */
    Object callback(int row, int col, Object value);
}
