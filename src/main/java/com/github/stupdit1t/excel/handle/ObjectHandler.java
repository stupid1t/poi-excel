package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;

import java.util.function.Function;


/**
 * 自定义转换
 *
 * @author 625
 */
public class ObjectHandler extends AbsCellVerifyRule<Object> {

    /**
     * 自定义验证
     *
     * @param allowNull    可为空
     * @param customVerify 自定义校验
     */
    public ObjectHandler(boolean allowNull, Function<Object, Object> customVerify) {
        super(allowNull, customVerify);
    }

    @Override
    public Object doHandle(String fieldName, Object cellValue) throws Exception {
        return customVerify.apply(cellValue);
    }
}
