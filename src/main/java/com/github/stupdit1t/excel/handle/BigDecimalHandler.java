package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;
import org.apache.commons.lang3.math.NumberUtils;

import java.math.BigDecimal;


/**
 * BigDecimal校验实体
 *
 * @author 625
 */
public class BigDecimalHandler extends BaseVerifyRule<BigDecimal> {

    /**
     * 常规验证
     *
     * @param allowNull 可为空
     */
    public BigDecimalHandler(boolean allowNull) {
        super(allowNull);
    }

    @Override
    public BigDecimal doHandle(String fieldName, String index, Object cellValue) throws Exception {
        String value = String.valueOf(cellValue);
        if (cellValue instanceof BigDecimal) {
            return (BigDecimal) cellValue;
        } else if (NumberUtils.isCreatable(value)) {
            return new BigDecimal(value);
        }
        throw PoiException.error(String.format(PoiConstant.INCORRECT_FORMAT_STR, fieldName, index));
    }
}
