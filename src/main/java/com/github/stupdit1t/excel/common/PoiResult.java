package com.github.stupdit1t.excel.common;

import java.util.Collections;
import java.util.List;

/**
 * excel 导入返回的实体类
 *
 * @param <T>
 * @author 625
 */
public class PoiResult<T> {

	private boolean success = true;

	private String message;

	private List<T> data;

	public boolean isSuccess() {
		return success;
	}

	public void setSuccess(boolean success) {
		this.success = success;
	}

	public String getMessage() {
		return message;
	}

	public void setMessage(String message) {
		this.message = message;
	}

	public List<T> getData() {
		return data;
	}

	public void setData(List<T> beans) {
		this.data = beans;
	}

	public static PoiResult fail() {
		PoiResult poiResult = new PoiResult();
		poiResult.setSuccess(false);
		poiResult.setMessage("读取Excel失败");
		poiResult.setData(Collections.emptyList());
		return poiResult;
	}

}
