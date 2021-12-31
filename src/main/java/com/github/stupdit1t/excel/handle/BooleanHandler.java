package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;

import java.util.function.Function;


/**
 * boolean校验实体
 *
 * @author 625
 */
public class BooleanHandler extends AbsCellVerifyRule<Boolean> {

    /**
     * 常规验证
     *
     * @param allowNull 可为空
     */
    public BooleanHandler(boolean allowNull) {
        super(allowNull);
    }

    /**
     * 自定义验证
     *
     * @param allowNull    可为空
     * @param customVerify 自定义校验
     */
    public BooleanHandler(boolean allowNull, Function<Object, Boolean> customVerify) {
        super(allowNull, customVerify);
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
