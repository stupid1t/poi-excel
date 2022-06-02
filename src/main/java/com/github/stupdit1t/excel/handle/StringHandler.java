package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;

import java.math.BigDecimal;
import java.util.regex.Pattern;


/**
 * 字符值校验实体
 *
 * @author 625
 */
public class StringHandler extends BaseVerifyRule<String> {

    /**
     * 正则验证
     */
    private String pattern;

    /**
     * 常规验证
     *
     * @param allowNull 是否可为空
     */
    public StringHandler(boolean allowNull) {
        super(allowNull);
    }

    /**
     * 常规验证
     *
     * @param allowNull 是否可为空
     */
    public StringHandler(boolean allowNull, String pattern) {
        super(allowNull);
        this.pattern = pattern;
    }


    @Override
    public String doHandle(String fieldName, String index, Object cellValue) throws Exception {
        String value = String.valueOf(cellValue);
        // 处理数值 转为 string包含E科学计数的问题
        if (cellValue instanceof Number) {
            value = new BigDecimal(value).toString();
        }
        if (pattern != null && !Pattern.matches(pattern, value)) {
            throw PoiException.error(String.format(PoiConstant.INCORRECT_FORMAT_STR, fieldName, index));
        }
        return value;
    }
}
