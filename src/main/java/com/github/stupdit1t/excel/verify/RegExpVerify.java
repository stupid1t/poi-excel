package com.github.stupdit1t.excel.verify;

import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.verify.rule.AbsCellVerifyRule;

import java.util.function.Function;
import java.util.regex.Pattern;


/**
 * 正则校验
 * 
 * @author 625
 *
 */
public class RegExpVerify extends AbsCellVerifyRule<String> {

	private String pattern;

	/**
	 * 常规验证
	 *
	 * @param allowNull
	 */
	public RegExpVerify(String pattern, boolean allowNull) {
		super(allowNull);
		this.pattern = pattern;
	}

	/**
	 * 自定义验证
	 *
	 * @param allowNull
	 * @param customVerify
	 */
	public RegExpVerify(String pattern, boolean allowNull, Function<Object, String> customVerify) {
		super(allowNull, customVerify);
		this.pattern = pattern;
	}

	@Override
	public String doVerify(String fieldName, Object cellValue) throws Exception{
		String value = String.valueOf(cellValue);
		if (!Pattern.matches(pattern, value)) {
			throw PoiException.error("格式不正确");
		}
		return value;
	}

}
