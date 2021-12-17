package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;
import org.apache.commons.lang3.math.NumberUtils;

import java.math.BigDecimal;
import java.util.function.Function;


/**
 * int数据校验
 *
 * @author 625
 */
public class IntegerHandler extends AbsCellVerifyRule<Integer> {
    /**
     * 常规验证
     *
     * @param allowNull
     */
    public IntegerHandler(boolean allowNull) {
        super(allowNull);
    }

    /**
     * 自定义验证
     *
     * @param allowNull
     * @param customVerify
     */
    public IntegerHandler(boolean allowNull, Function<Object, Integer> customVerify) {
        super(allowNull, customVerify);
    }

    @Override
    public Integer doHandle(String fieldName, Object cellValue) throws Exception {
        String value = String.valueOf(cellValue);
        if (cellValue instanceof Integer) {
            return (Integer) cellValue;
        } else if (NumberUtils.isNumber(value)) {
            return new BigDecimal(value).intValue();
        }
        throw PoiException.error(fieldName+"格式不正确");
    }
}
