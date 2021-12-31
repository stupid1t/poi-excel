package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;
import org.apache.commons.lang3.math.NumberUtils;

import java.math.BigDecimal;
import java.util.function.Function;


/**
 * long校验实体
 * 
 * @author 625
 *
 */
public class LongHandler extends AbsCellVerifyRule<Long> {
	/**
	 * 常规验证
	 *
	 * @param allowNull 可为空
	 */
	public LongHandler(boolean allowNull) {
		super(allowNull);
	}

	/**
	 * 自定义验证
	 *
	 * @param allowNull 可为空
	 * @param customVerify 自定义校验
	 */
	public LongHandler(boolean allowNull, Function<Object, Long> customVerify) {
		super(allowNull, customVerify);
	}

	@Override
	public Long doHandle(String fieldName, Object cellValue) throws Exception {
		String value = String.valueOf(cellValue);
		if (cellValue instanceof Long) {
			return (Long) cellValue;
		} else if (NumberUtils.isNumber(value)) {
			return new BigDecimal(value).longValue();
		}
		throw PoiException.error(fieldName+"格式不正确");
	}
}
