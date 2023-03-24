package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.core.parse.OpsColumn;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;


/**
 * boolean校验实体
 *
 * @author 625
 */
public class BooleanHandler<R> extends BaseVerifyRule<Boolean, R> {

	/**
	 * 常规验证
	 *
	 * @param allowNull 可为空
	 */
	public BooleanHandler(boolean allowNull, OpsColumn<R> opsColumn) {
		super(allowNull, opsColumn);
	}

	@Override
	public Boolean doHandle(Object cellValue) throws Exception {
		if (cellValue instanceof Boolean) {
			return (Boolean) cellValue;
		} else {
			String value = String.valueOf(cellValue);
			if (this.trim) {
				value = value.trim();
			}
			return Boolean.parseBoolean(value);
		}
	}
}
