package com.github.stupdit1t.excel.handle.rule;

import com.github.stupdit1t.excel.common.PoiException;
import org.apache.commons.lang3.StringUtils;

import java.util.function.Function;

/**
 * 列校验和格式化接口
 *
 * @author 625
 */
public abstract class AbsCellVerifyRule<T> {

    /**
     * 是否可为空
     */
    protected boolean allowNull;

    /**
     * 自定义校验, 覆盖摸默认校验
     */
    protected Function<Object, T> customVerify;

    /**
     * 构建校验规则
     *
     * @param allowNull 是否为空
     */
    public AbsCellVerifyRule(boolean allowNull) {
        this.allowNull = allowNull;
    }

    /**
     * 构建校验规则
     *
     * @param allowNull    是否为空
     * @param customVerify 自定义校验
     */
    public AbsCellVerifyRule(boolean allowNull, Function<Object, T> customVerify) {
        this.customVerify = customVerify;
        this.allowNull = allowNull;
    }

    /**
     * 判空处理
     *
     * @param fieldName 列名称
     * @param value     列值
     * @throws PoiException
     */
    public Object handleNull(String fieldName, Object value) throws PoiException {
        if (value == null || StringUtils.isBlank(String.valueOf(value))) {
            if (this.allowNull) {
                return null;
            } else {
                throw PoiException.error(fieldName + "不能为空");
            }
        }
        return value;
    }

    /**
     * 校验单元格值
     *
     * @param fieldName 列名称
     * @param cellValue 列值
     * @throws Exception
     */
    public T handle(String filed, String fieldName, Object cellValue) throws PoiException {
        // 空值处理
        cellValue = handleNull(fieldName, cellValue);
        if (cellValue == null) {
            return null;
        }
        T endVal;
        if (null != customVerify) {
            endVal = customVerify.apply(cellValue);
        } else {
            try {
                endVal = doHandle(filed, cellValue);
            } catch (PoiException e) {
                throw e;
            } catch (Exception e) {
                throw PoiException.error(fieldName + "格式不正确");
            }
        }
        return endVal;
    }

    /**
     * 校验单元格值
     *
     * @param fieldName 列名称
     * @param cellValue 列值
     * @throws Exception
     */
    public abstract T doHandle(String fieldName, Object cellValue) throws Exception;
}
