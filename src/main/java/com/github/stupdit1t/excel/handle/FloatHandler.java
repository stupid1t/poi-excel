package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.core.parse.OpsColumn;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;
import org.apache.commons.lang3.math.NumberUtils;

import java.math.BigDecimal;


/**
 * float校验实体
 *
 * @author 625
 */
public class FloatHandler<R> extends BaseVerifyRule<Float, R> {
    /**
     * 常规验证
     *
     * @param allowNull 可为空
     */
    public FloatHandler(boolean allowNull, OpsColumn<R> opsColumn) {
        super(allowNull, opsColumn);
    }

    @Override
    protected Float doHandle(int row, int col, Object cellValue) throws Exception {
        if (cellValue instanceof Float) {
            return (Float) cellValue;
        } else {
            String value = String.valueOf(cellValue);
            if (this.trim) {
                value = value.trim();
            }
            if (NumberUtils.isCreatable(value)) {
                return new BigDecimal(value).floatValue();
            }
        }
        throw PoiException.error(PoiConstant.INCORRECT_FORMAT_STR);
    }
}
