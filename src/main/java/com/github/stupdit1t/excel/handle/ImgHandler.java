package com.github.stupdit1t.excel.handle;


import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;

/**
 * 图片校验实体
 *
 * @author 625
 */
public class ImgHandler extends BaseVerifyRule<byte[]> {

    /**
     * 常规验证
     *
     * @param allowNull 可为空
     */
    public ImgHandler(boolean allowNull) {
        super(allowNull);
    }

    @Override
    public byte[] doHandle(String fieldName, String index, Object cellValue) throws Exception {
        if (cellValue instanceof byte[]) {
            return (byte[]) cellValue;
        }
        throw PoiException.error(String.format(PoiConstant.INCORRECT_FORMAT_STR, fieldName, index));
    }

}
