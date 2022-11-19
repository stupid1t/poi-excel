package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;
import org.apache.commons.lang3.math.NumberUtils;

import java.math.BigDecimal;


/**
 * short校验实体
 *
 * @author 625
 */
public class ShortHandler extends BaseVerifyRule<Short> {
    /**
     * 常规验证
     *
     * @param allowNull 可为空
     */
    public ShortHandler(boolean allowNull) {
        super(allowNull);
    }

    @Override
    public Short doHandle(String fieldName, String index, Object cellValue) throws Exception {
        if (cellValue instanceof Short) {
            return (Short) cellValue;
        } else {
            String value = String.valueOf(cellValue);
            if (this.trim) {
                value = value.trim();
            }
            if (NumberUtils.isCreatable(value)) {
                return new BigDecimal(value).shortValue();
            }
        }
        throw PoiException.error(String.format(PoiConstant.INCORRECT_FORMAT_STR, fieldName, index));
    }
}
