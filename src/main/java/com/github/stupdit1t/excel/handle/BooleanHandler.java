package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;

import java.util.function.Function;


/**
 * boolean校验实体
 *
 * @author 625
 */
public class BooleanHandler extends AbsCellVerifyRule<Boolean> {

    private String pattern;

    /**
     * 常规验证
     *
     * @param allowNull
     */
    public BooleanHandler(boolean allowNull) {
        super(allowNull);
    }

    /**
     * 自定义验证
     *
     * @param allowNull
     * @param customVerify
     */
    public BooleanHandler(boolean allowNull, Function<Object, Boolean> customVerify) {
        super(allowNull, customVerify);
    }

    /**
     * 常规验证
     *
     * @param allowNull
     */
    public BooleanHandler(String pattern, boolean allowNull) {
        super(allowNull);
        this.pattern = pattern;
    }

    /**
     * 自定义验证
     *
     * @param allowNull
     * @param customVerify
     */
    public BooleanHandler(String pattern, boolean allowNull, Function<Object, Boolean> customVerify) {
        super(allowNull, customVerify);
        this.pattern = pattern;
    }

    @Override
    public Boolean doHandle(String fieldName, Object cellValue) throws Exception {
        if (cellValue instanceof Boolean) {
            return (Boolean) cellValue;
        } else {
            return Boolean.parseBoolean(String.valueOf(cellValue));
        }
    }
}
