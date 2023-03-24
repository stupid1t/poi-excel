package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.core.parse.OpsColumn;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;

import java.math.BigDecimal;
import java.util.regex.Pattern;


/**
 * 字符值校验实体
 *
 * @author 625
 */
public class StringHandler<R> extends BaseVerifyRule<String, R> {

	/**
	 * 正则验证
	 */
	private String pattern;

	/**
	 * 常规验证
	 *
	 * @param allowNull 可为空
	 */
	public StringHandler(boolean allowNull, OpsColumn<R> opsColumn) {
		super(allowNull, opsColumn);
	}

	/**
	 * 格式
	 *
	 * @param pattern 是否可为空
	 */
	public StringHandler<R> pattern(String pattern) {
		this.pattern = pattern;
		return this;
	}


	@Override
	public String doHandle(int row, int col, Object cellValue) throws Exception {
		String value = String.valueOf(cellValue);
		if (this.trim) {
			value = value.trim();
		}
		// 处理数值 转为 string包含E科学计数的问题
		if (cellValue instanceof Number) {
			value = new BigDecimal(value).toString();
		}
		if (pattern != null && !Pattern.matches(pattern, value)) {
			throw PoiException.error(PoiConstant.INCORRECT_FORMAT_STR);
		}
		return value;
	}
}
