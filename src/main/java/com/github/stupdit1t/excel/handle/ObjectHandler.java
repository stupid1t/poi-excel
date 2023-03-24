package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.callback.CellCallback;
import com.github.stupdit1t.excel.core.parse.OpsColumn;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;


/**
 * 自定义转换
 *
 * @author 625
 */
public class ObjectHandler<R> extends BaseVerifyRule<Object, R> {

    private CellCallback doHandleSub;

    /**
     * 自定义验证
     *
     * @param allowNull 可为空
     */
    public ObjectHandler(boolean allowNull, OpsColumn<R> opsColumn, CellCallback handle) {
        super(allowNull, opsColumn);
        this.doHandleSub = handle;
    }

    @Override
    public Object doHandle(int row, int col, Object cellValue) throws Exception {
        return doHandleSub.callback(row, col, cellValue);
    }
}
