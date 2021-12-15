package com.github.stupdit1t.excel.verify;

import com.github.stupdit1t.excel.verify.rule.AbsCellVerifyRule;

import java.util.function.Function;


/**
 * double校验实体
 * 
 * @author 625
 *
 */
public class DoubleVerify extends AbsCellVerifyRule<Double> {

	/**
	 * 常规验证
	 *
	 * @param allowNull
	 */
	public DoubleVerify(boolean allowNull) {
		super(allowNull);
	}

	/**
	 * 自定义验证
	 *
	 * @param allowNull
	 * @param customVerify
	 */
	public DoubleVerify(boolean allowNull, Function<Object, Double> customVerify) {
		super(allowNull, customVerify);
	}

	@Override
	public Double doVerify(String fieldName, Object cellValue) throws Exception {
		if (cellValue instanceof Double) {
			return (Double) cellValue;
		} else if (cellValue instanceof Number) {
			Number old = (Number) cellValue;
			return old.doubleValue();
		}
		return Double.parseDouble(cellValue.toString());
	}

}
