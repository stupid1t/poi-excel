package com.github.stupdit1t.excel.verify;


import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.verify.rule.AbsCellVerifyRule;

import java.util.function.Function;

/**
 * 字符值校验实体
 * 
 * @author 625
 *
 */
public class ImgVerify extends AbsCellVerifyRule<byte[]> {

	/**
	 * 常规验证
	 *
	 * @param allowNull
	 */
	public ImgVerify(boolean allowNull) {
		super(allowNull);
	}

	/**
	 * 自定义验证
	 *
	 * @param allowNull
	 * @param customVerify
	 */
	public ImgVerify(boolean allowNull, Function<Object, byte[]> customVerify) {
		super(allowNull, customVerify);
	}

	@Override
	public byte[] doVerify(String fieldName, Object cellValue) throws Exception {
		if (cellValue instanceof byte[]) {
			return (byte[]) cellValue;
		}
		throw PoiException.error("图片验证未知异常!");
	}

}
