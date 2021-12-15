package com.github.stupdit1t.excel.verify;

import com.github.stupdit1t.excel.verify.rule.AbsCellVerifyRule;

import java.math.BigDecimal;
import java.util.function.Function;


/**
 * BigDecimal校验实体
 *
 * @author 625
 */
public class BigDecimalVerify extends AbsCellVerifyRule<BigDecimal> {

    /**
     * 常规验证
     *
     * @param allowNull
     */
    public BigDecimalVerify(boolean allowNull) {
        super(allowNull);
    }

    /**
     * 自定义验证
     *
     * @param allowNull
     * @param customVerify
     */
    public BigDecimalVerify(boolean allowNull, Function<Object, BigDecimal> customVerify) {
        super(allowNull, customVerify);
    }

    @Override
    public BigDecimal doVerify(String fieldName, Object cellValue) throws Exception {
        if (cellValue instanceof BigDecimal) {
            return (BigDecimal) cellValue;
        } else if (cellValue instanceof Number) {
            Number old = (Number) cellValue;
            return new BigDecimal(old.toString());
        }
        return new BigDecimal(cellValue.toString());
    }
}
