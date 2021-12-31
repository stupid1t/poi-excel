package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;

import java.util.function.Function;


/**
 * char校验实体
 *
 * @author 625
 */
public class CharHandler extends AbsCellVerifyRule<Character> {

    /**
     * 常规验证
     *
     * @param allowNull 可为空
     */
    public CharHandler(boolean allowNull) {
        super(allowNull);
    }

    /**
     * 自定义验证
     *
     * @param allowNull    可为空
     * @param customVerify 自定义校验
     */
    public CharHandler(boolean allowNull, Function<Object, Character> customVerify) {
        super(allowNull, customVerify);
    }


    @Override
    public Character doHandle(String fieldName, Object cellValue) throws Exception {
        if (cellValue instanceof Character) {
            return (Character) cellValue;
        } else {
            return String.valueOf(cellValue).toCharArray()[0];
        }
    }
}
