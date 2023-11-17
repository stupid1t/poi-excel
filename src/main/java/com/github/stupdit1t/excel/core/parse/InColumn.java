package com.github.stupdit1t.excel.core.parse;

import com.github.stupdit1t.excel.core.AbsParent;
import com.github.stupdit1t.excel.handler.BaseVerifyRule;

/**
 * 列的定义
 *
 * @author 625
 */
public class InColumn<R> extends AbsParent<OpsColumn<R>> {

    /**
     * 导入下标
     */
    final String index;

    /**
     * 字段
     */
    final String field;

    /**
     * 验证规则
     */
    BaseVerifyRule<R> cellVerifyRule = new BaseVerifyRule<>(this.parent);

    public InColumn(OpsColumn<R> opsColumn, String index, String field) {
        super(opsColumn);
        this.index = index;
        this.field = field;
    }

    /**
     * 获取字段
     *
     * @return String
     */
    public String getField() {
        return field;
    }

    /**
     * 获取校验规则
     *
     * @return BaseVerifyRule
     */
    public BaseVerifyRule<R> getCellVerifyRule() {
        return cellVerifyRule;
    }
}
