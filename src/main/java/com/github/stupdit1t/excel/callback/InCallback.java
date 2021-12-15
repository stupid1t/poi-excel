package com.github.stupdit1t.excel.callback;

/**
 * 导入回调函数
 *
 * @author 625
 */

@FunctionalInterface
public interface InCallback<T> {
	/**
	 * 导入回调
	 * 
	 * @param t 当前行数据
	 * @param rowNum 当前行号
	 * @throws Exception
	 */
	void callback(T t, int rowNum) throws Exception;
}
