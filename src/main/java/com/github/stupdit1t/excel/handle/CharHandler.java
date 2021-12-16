package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;

import java.util.function.Function;


/**
 * char校验实体
 *
 * @author 625
 */
public class CharHandler extends AbsCellVerifyRule<Character> {

    private String pattern;

    /**
     * 常规验证
     *
     * @param allowNull
     */
    public CharHandler(boolean allowNull) {
        super(allowNull);
    }

    /**
     * 自定义验证
     *
     * @param allowNull
     * @param customVerify
     */
    public CharHandler(boolean allowNull, Function<Object, Character> customVerify) {
        super(allowNull, customVerify);
    }

    /**
     * 常规验证
     *
     * @param allowNull
     */
    public CharHandler(String pattern, boolean allowNull) {
        super(allowNull);
        this.pattern = pattern;
    }

    /**
     * 自定义验证
     *
     * @param allowNull
     * @param customVerify
     */
    public CharHandler(String pattern, boolean allowNull, Function<Object, Character> customVerify) {
        super(allowNull, customVerify);
        this.pattern = pattern;
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
