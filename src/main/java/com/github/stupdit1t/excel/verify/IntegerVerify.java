package com.github.stupdit1t.excel.verify;

import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.verify.rule.AbsCellVerifyRule;

import java.util.function.Function;


/**
 * int数据校验
 *
 * @author 625
 */
public class IntegerVerify extends AbsCellVerifyRule<Integer> {
    /**
     * 常规验证
     *
     * @param allowNull
     */
    public IntegerVerify(boolean allowNull) {
        super(allowNull);
    }

    /**
     * 自定义验证
     *
     * @param allowNull
     * @param customVerify
     */
    public IntegerVerify(boolean allowNull, Function<Object, Integer> customVerify) {
        super(allowNull, customVerify);
    }

    @Override
    public Integer doVerify(String fieldName, Object cellValue) throws Exception {
        if (cellValue instanceof Integer) {
            return (Integer) cellValue;
        } else if (cellValue instanceof Number) {
            Number old = (Number) cellValue;
            double diff = old.doubleValue() - old.intValue();
            if (diff > 0) {
                throw PoiException.error("格式不正确");
            }
            return old.intValue();
        }
        return Integer.parseInt(cellValue.toString());
    }
}
