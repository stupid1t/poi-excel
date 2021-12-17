package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;
import org.apache.commons.lang3.math.NumberUtils;

import java.math.BigDecimal;
import java.util.function.Function;


/**
 * double校验实体
 * 
 * @author 625
 *
 */
public class DoubleHandler extends AbsCellVerifyRule<Double> {

	/**
	 * 常规验证
	 *
	 * @param allowNull
	 */
	public DoubleHandler(boolean allowNull) {
		super(allowNull);
	}

	/**
	 * 自定义验证
	 *
	 * @param allowNull
	 * @param customVerify
	 */
	public DoubleHandler(boolean allowNull, Function<Object, Double> customVerify) {
		super(allowNull, customVerify);
	}

	@Override
	public Double doHandle(String fieldName, Object cellValue) throws Exception {
		String value = String.valueOf(cellValue);
		if (cellValue instanceof Double) {
			return (Double) cellValue;
		} else if (NumberUtils.isNumber(value)) {
			return new BigDecimal(value).doubleValue();
		}
		throw PoiException.error(fieldName+"格式不正确");
	}

}