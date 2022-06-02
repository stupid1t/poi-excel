package com.github.stupdit1t.excel.callback;

/**
 * 导入回调函数
 *
 * @author 625
 */

@FunctionalInterface
public interface InCallback<R> {
    /**
     * 导入回调
     *
     * @param row    当前数据
     * @param rowNum 行号
     */
    void callback(R row, int rowNum) throws Exception;
}
