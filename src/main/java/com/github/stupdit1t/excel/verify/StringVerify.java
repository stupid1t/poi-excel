package com.github.stupdit1t.excel.verify;

import com.github.stupdit1t.excel.verify.rule.AbsCellVerifyRule;

import java.util.function.Function;


/**
 * 字符值校验实体
 * 
 * @author 625
 *
 */
public class StringVerify extends AbsCellVerifyRule<String> {
	/**
	 * 常规验证
	 *
	 * @param allowNull
	 */
	public StringVerify(boolean allowNull) {
		super(allowNull);
	}

	/**
	 * 自定义验证
	 *
	 * @param allowNull
	 * @param customVerify
	 */
	public StringVerify(boolean allowNull, Function<Object, String> customVerify) {
		super(allowNull, customVerify);
	}

	@Override
	public String doVerify(String fieldName, Object cellValue) throws Exception {
		return String.valueOf(cellValue);
	}
}
