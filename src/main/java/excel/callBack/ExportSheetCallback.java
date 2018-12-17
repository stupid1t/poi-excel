package excel.callBack;

import excel.ExcelUtils.Column;

public interface ExportSheetCallback<T> {
	/**
	 * 导出回调
	 * 
	 * @param fieldName 导出字段名
	 * @param value     当前单元格值
	 * @param row       当前实体
	 * @param col       当前单元格样式
	 * @return 返回重置后的单元格值
	 * @throws Exception
	 */
	Object callback(String fieldName, Object value, T t, Column customStyle);
}
