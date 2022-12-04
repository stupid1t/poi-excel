package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.core.parse.OpsColumn;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;


/**
 * char校验实体
 *
 * @author 625
 */
public class CharHandler<R> extends BaseVerifyRule<Character, R> {

    /**
     * 常规验证
     *
     * @param allowNull 可为空
     */
    public CharHandler(boolean allowNull, OpsColumn<R> opsColumn) {
        super(allowNull, opsColumn);
    }

    @Override
    public Character doHandle(String fieldName, String index, Object cellValue) throws Exception {
        if (cellValue instanceof Character) {
            return (Character) cellValue;
        } else {
            String value = String.valueOf(cellValue);
            if (this.trim) {
                value = value.trim();
            }
            return value.toCharArray()[0];
        }
    }
}
