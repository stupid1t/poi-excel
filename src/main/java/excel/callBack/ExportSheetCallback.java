package excel.callBack;

public interface ExportSheetCallback<T> {
	/**
	 * 导出回调
	 * 
	 * @param fieldName 导出字段名
	 * @param value 当前单元格值
	 * @param style 当前数据行值
	 * @return 返回重置后的单元格值
	 * @throws Exception
	 */
	Object callback(String fieldName, Object value, T t) throws Exception;
}
