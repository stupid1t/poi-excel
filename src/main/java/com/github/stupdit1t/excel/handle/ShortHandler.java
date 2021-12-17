package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;
import org.apache.commons.lang3.math.NumberUtils;

import java.math.BigDecimal;
import java.util.function.Function;


/**
 * short校验实体
 * 
 * @author 625
 *
 */
public class ShortHandler extends AbsCellVerifyRule<Short> {
	/**
	 * 常规验证
	 *
	 * @param allowNull
	 */
	public ShortHandler(boolean allowNull) {
		super(allowNull);
	}

	/**
	 * 自定义验证
	 *
	 * @param allowNull
	 * @param customVerify
	 */
	public ShortHandler(boolean allowNull, Function<Object, Short> customVerify) {
		super(allowNull, customVerify);
	}

	@Override
	public Short doHandle(String fieldName, Object cellValue) throws Exception {
		String value = String.valueOf(cellValue);
		if (cellValue instanceof Short) {
			return (Short) cellValue;
		} else if (NumberUtils.isNumber(value)) {
			return new BigDecimal(value).shortValue();
		}
		throw PoiException.error(fieldName+"格式不正确");
	}
}
