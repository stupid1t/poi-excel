package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;


/**
 * boolean校验实体
 *
 * @author 625
 */
public class BooleanHandler extends BaseVerifyRule<Boolean> {

    /**
     * 常规验证
     *
     * @param allowNull 可为空
     */
    public BooleanHandler(boolean allowNull) {
        super(allowNull);
    }

    @Override
    public Boolean doHandle(String fieldName, String index, Object cellValue) throws Exception {
        if (cellValue instanceof Boolean) {
            return (Boolean) cellValue;
        } else {
            return Boolean.parseBoolean(String.valueOf(cellValue));
        }
    }
}
