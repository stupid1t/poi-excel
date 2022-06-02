package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;


/**
 * char校验实体
 *
 * @author 625
 */
public class CharHandler extends BaseVerifyRule<Character> {

    /**
     * 常规验证
     *
     * @param allowNull 可为空
     */
    public CharHandler(boolean allowNull) {
        super(allowNull);
    }

    @Override
    public Character doHandle(String fieldName, String index, Object cellValue) throws Exception {
        if (cellValue instanceof Character) {
            return (Character) cellValue;
        } else {
            return String.valueOf(cellValue).toCharArray()[0];
        }
    }
}
