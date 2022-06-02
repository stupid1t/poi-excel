package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;
import org.apache.commons.lang3.math.NumberUtils;

import java.math.BigDecimal;


/**
 * long校验实体
 *
 * @author 625
 */
public class LongHandler extends BaseVerifyRule<Long> {
    /**
     * 常规验证
     *
     * @param allowNull 可为空
     */
    public LongHandler(boolean allowNull) {
        super(allowNull);
    }

    @Override
    public Long doHandle(String fieldName, String index, Object cellValue) throws Exception {
        String value = String.valueOf(cellValue);
        if (cellValue instanceof Long) {
            return (Long) cellValue;
        } else if (NumberUtils.isCreatable(value)) {
            return new BigDecimal(value).longValue();
        }
        throw PoiException.error(String.format(PoiConstant.INCORRECT_FORMAT_STR, fieldName, index));
    }
}
