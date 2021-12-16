package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;
import org.apache.commons.lang3.math.NumberUtils;

import java.math.BigDecimal;
import java.util.function.Function;


/**
 * BigDecimal校验实体
 *
 * @author 625
 */
public class BigDecimalHandler extends AbsCellVerifyRule<BigDecimal> {

    /**
     * 常规验证
     *
     * @param allowNull
     */
    public BigDecimalHandler(boolean allowNull) {
        super(allowNull);
    }

    /**
     * 自定义验证
     *
     * @param allowNull
     * @param customVerify
     */
    public BigDecimalHandler(boolean allowNull, Function<Object, BigDecimal> customVerify) {
        super(allowNull, customVerify);
    }

    @Override
    public BigDecimal doHandle(String fieldName, Object cellValue) throws Exception {
        String value = String.valueOf(cellValue);
        if (cellValue instanceof BigDecimal) {
            return (BigDecimal) cellValue;
        } else if (NumberUtils.isNumber(value)) {
            return new BigDecimal(value);
        }
        throw PoiException.error("格式不支持");
    }
}
