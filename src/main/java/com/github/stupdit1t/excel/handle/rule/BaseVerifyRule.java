package com.github.stupdit1t.excel.handle.rule;

import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiException;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
 * 列校验和格式化接口
 *
 * @author 625
 */
public abstract class BaseVerifyRule<T> {

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
     * 构建校验规则
     *
     * @param allowNull 是否为空
     */
    public BaseVerifyRule(boolean allowNull) {
        this.allowNull = allowNull;
    }

    /**
     * 判空处理
     *
     * @param fieldName 列名称
     * @param value     列值
     */
    public Object handleNull(String fieldName, String index, Object value) throws PoiException {
        if (ObjectUtils.isEmpty(value)) {
            if (this.allowNull) {
                return null;
            } else {
                throw PoiException.error(String.format(PoiConstant.NOT_EMPTY_STR, fieldName, index));
            }
        }
        return value;
    }

    /**
     * 校验单元格值
     *
     * @param fieldName 列名称
     * @param cellValue 列值
     */
    public T handle(String fieldName, String index, Object cellValue) throws PoiException {
        // 空值处理
        cellValue = handleNull(fieldName, index, cellValue);
        if (cellValue == null) {
            return null;
        }
        T endVal;
        try {
            endVal = doHandle(fieldName, index, cellValue);
        } catch (PoiException e) {
            throw e;
        } catch (Exception e) {
            LOG.error(e);
            throw PoiException.error(String.format(PoiConstant.INCORRECT_FORMAT_STR, fieldName, index));
        }
        return endVal;
    }

    /**
     * 校验单元格值
     *
     * @param fieldName 列名称
     * @param cellValue 列值
     */
    public abstract T doHandle(String fieldName, String index, Object cellValue) throws Exception;

    /**
     * 设置是否可为空
     *
     * @param allowNull 可为空
     */
    public void setAllowNull(boolean allowNull) {
        this.allowNull = allowNull;
    }

    /**
     * 是否去空格
     * @param trim 是 去除两边空格 否  不去除
     */
    public void setTrim(boolean trim) {
        this.trim = trim;
    }

}
