package com.github.stupdit1t.excel.handle;


import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;

import java.util.function.Function;

/**
 * 图片校验实体
 *
 * @author 625
 */
public class ImgHandler extends AbsCellVerifyRule<byte[]> {

    /**
     * 常规验证
     *
     * @param allowNull
     */
    public ImgHandler(boolean allowNull) {
        super(allowNull);
    }

    /**
     * 自定义验证
     *
     * @param allowNull
     * @param customVerify
     */
    public ImgHandler(boolean allowNull, Function<Object, byte[]> customVerify) {
        super(allowNull, customVerify);
    }

    @Override
    public byte[] doHandle(String fieldName, Object cellValue) throws Exception {
        if (cellValue instanceof byte[]) {
            return (byte[]) cellValue;
        }
        throw PoiException.error(fieldName + "请检查图片数据格式");
    }

}
