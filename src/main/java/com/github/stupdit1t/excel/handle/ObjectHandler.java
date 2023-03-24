package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.core.parse.OpsColumn;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;

import java.util.function.Function;


/**
 * 自定义转换
 *
 * @author 625
 */
public class ObjectHandler<R> extends BaseVerifyRule<Object, R> {

    private Function<Object, Object> doHandleSub;

    /**
     * 自定义验证
     *
     * @param allowNull 可为空
     */
    public ObjectHandler(boolean allowNull, OpsColumn<R> opsColumn, Function<Object, Object> handle) {
        super(allowNull, opsColumn);
        this.doHandleSub = handle;
    }

    @Override
    public Object doHandle(Object cellValue) throws Exception {
        return doHandleSub.apply(cellValue);
    }
}
