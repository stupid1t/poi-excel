package com.github.stupdit1t.excel.callback;


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
	 * @param value       当前单元格值
	 * @param t           当前实体
	 * @param customStyle 自定义单元格样式
	 * @return 返回重置后的单元格值
	 */
	Object callback(Object value, R t);
}
