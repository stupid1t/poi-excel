package com.github.stupdit1t.excel.handle;


import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.core.parse.OpsColumn;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;

/**
 * 图片校验实体
 *
 * @author 625
 */
public class ImgHandler<R> extends BaseVerifyRule<byte[], R> {

    /**
     * 常规验证
     *
     * @param allowNull 可为空
     */
    public ImgHandler(boolean allowNull, OpsColumn<R> opsColumn) {
        super(allowNull, opsColumn);
    }

    @Override
    public byte[] doHandle(int row, int col, Object cellValue) throws Exception {
        if (cellValue instanceof byte[]) {
            return (byte[]) cellValue;
        }
        throw PoiException.error(PoiConstant.INCORRECT_FORMAT_STR);
    }

}
