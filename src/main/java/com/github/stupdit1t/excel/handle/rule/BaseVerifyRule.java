package com.github.stupdit1t.excel.handle.rule;

import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.core.AbsParent;
import com.github.stupdit1t.excel.core.parse.OpsColumn;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.function.Function;

/**
 * 列校验和格式化接口
 *
 * @author 625
 */
public abstract class BaseVerifyRule<T, R> extends AbsParent<OpsColumn<R>> {

    private static final Logger LOG = LogManager.getLogger(BaseVerifyRule.class);

    /**
     * 是否可为空
     */
    protected boolean allowNull;

    /**
     * 是否去空格
     */
    protected boolean trim;

    /**
     * 默认值
     */
    protected T defaultValue;

    /**
     * 映射转换
     */
    private Function<T, T> mapping;

    /**
     * 构建校验规则
     *
     * @param allowNull 是否为空
     */
    public BaseVerifyRule(boolean allowNull, OpsColumn<R> parent) {
        super(parent);
        this.allowNull = allowNull;
    }

    /**
     * 判空处理
     *
     * @param value 列值
     */
    protected Object handleNull(Object value) throws PoiException {
        if (ObjectUtils.isEmpty(value)) {
            if (this.allowNull) {
                return null;
            } else {
                throw PoiException.error(PoiConstant.NOT_EMPTY_STR);
            }
        }
        return value;
    }

    /**
     * 校验单元格值
     *
     * @param cellValue 列值
     */
    public T handle(int row, int col, Object cellValue) throws Exception {
        // 空值处理
        cellValue = handleNull(cellValue);
        if (ObjectUtils.isEmpty(cellValue)) {
            return this.defaultValue;
        }
        T data = doHandle(row, col, cellValue);

        // 数据映射转换，也可做判断
        if (mapping != null) {
            return mapping.apply(data);
        }
        return data;
    }

    /**
     * 不能为空
     *
     * @return InColumn<R>
     */
    public BaseVerifyRule<T, R> notNull() {
        this.allowNull = false;
        return this;
    }

    /**
     * 去除两边空格
     *
     * @return InColumn<R>
     */
    public BaseVerifyRule<T, R> trim() {
        this.trim = true;
        return this;
    }

    /**
     * 去除两边空格
     *
     * @return InColumn<R>
     */
    public BaseVerifyRule<T, R> defaultValue(T defaultValue) {
        this.defaultValue = defaultValue;
        return this;
    }


    /**
     * 转换or映射or判断
     */
    public BaseVerifyRule<T, R> map(Function<T, T> mapping) {
        this.mapping = mapping;
        return this;
    }

    /**
     * 校验单元格值
     *
     * @param row       行号
     * @param col       列号
     * @param cellValue 列值
     */
    protected abstract T doHandle(int row, int col, Object cellValue) throws Exception;
}
