package com.github.stupdit1t.excel.verify;

import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.verify.rule.AbsCellVerifyRule;

import java.util.function.Function;


/**
 * long校验实体
 * 
 * @author 625
 *
 */
public class LongVerify extends AbsCellVerifyRule<Long> {
	/**
	 * 常规验证
	 *
	 * @param allowNull
	 */
	public LongVerify(boolean allowNull) {
		super(allowNull);
	}

	/**
	 * 自定义验证
	 *
	 * @param allowNull
	 * @param customVerify
	 */
	public LongVerify(boolean allowNull, Function<Object, Long> customVerify) {
		super(allowNull, customVerify);
	}

	@Override
	public Long doVerify(String fieldName, Object cellValue) throws Exception {
		if (cellValue instanceof Long) {
			return (Long) cellValue;
		} else if (cellValue instanceof Number) {
			Number old = (Number) cellValue;
			double diff = old.doubleValue() - old.longValue();
			if (diff > 0) {
				throw PoiException.error("格式不正确");
			}
			return old.longValue();
		}
		return Long.parseLong(cellValue.toString());
	}
}
