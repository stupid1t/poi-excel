package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;
import org.apache.commons.lang3.math.NumberUtils;

import java.math.BigDecimal;
import java.util.function.Function;


/**
 * float校验实体
 * 
 * @author 625
 *
 */
public class FloatHandler extends AbsCellVerifyRule<Float> {
	/**
	 * 常规验证
	 *
	 * @param allowNull
	 */
	public FloatHandler(boolean allowNull) {
		super(allowNull);
	}

	/**
	 * 自定义验证
	 *
	 * @param allowNull
	 * @param customVerify
	 */
	public FloatHandler(boolean allowNull, Function<Object, Float> customVerify) {
		super(allowNull, customVerify);
	}

	@Override
	public Float doHandle(String fieldName, Object cellValue) throws Exception {
		String value = String.valueOf(cellValue);
		if (cellValue instanceof Float) {
			return (Float) cellValue;
		} else if (NumberUtils.isNumber(value)) {
			return new BigDecimal(value).floatValue();
		}
		throw PoiException.error(fieldName+"格式不正确");
	}
}
