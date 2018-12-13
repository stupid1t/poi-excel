package excel.callBack;

public interface ParseSheetCallback<T> {
	/**
	 * 导入回调
	 * 
	 * @param t 当前行数据
	 * @param rowNum 当前行号
	 * @throws Exception
	 */
	void callback(T t, int rowNum) throws Exception;
}
