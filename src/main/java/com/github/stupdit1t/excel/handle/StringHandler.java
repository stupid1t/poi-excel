package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;

import java.util.function.Function;
import java.util.regex.Pattern;


/**
 * 字符值校验实体
 *
 * @author 625
 */
public class StringHandler extends AbsCellVerifyRule<String> {

    private String pattern;

    /**
     * 常规验证
     *
     * @param allowNull
     */
    public StringHandler(boolean allowNull) {
        super(allowNull);
    }

    /**
     * 自定义验证
     *
     * @param allowNull
     * @param customVerify
     */
    public StringHandler(boolean allowNull, Function<Object, String> customVerify) {
        super(allowNull, customVerify);
    }

    /**
     * 常规验证
     *
     * @param allowNull
     */
    public StringHandler(String pattern, boolean allowNull) {
        super(allowNull);
        this.pattern = pattern;
    }

    /**
     * 自定义验证
     *
     * @param allowNull
     * @param customVerify
     */
    public StringHandler(String pattern, boolean allowNull, Function<Object, String> customVerify) {
        super(allowNull, customVerify);
        this.pattern = pattern;
    }

    @Override
    public String doHandle(String fieldName, Object cellValue) throws Exception {
        String value = String.valueOf(cellValue);
        if (pattern != null && !Pattern.matches(pattern, value)) {
            throw PoiException.error("格式不正确");
        }
        return value;
    }
}
