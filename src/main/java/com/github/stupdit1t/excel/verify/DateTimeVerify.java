package com.github.stupdit1t.excel.verify;

import com.github.stupdit1t.excel.verify.rule.AbsCellVerifyRule;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.function.Function;


/**
 * 日期校验实体
 *
 * @author 625
 */
public class DateTimeVerify extends AbsCellVerifyRule<Date> {

    private String pattern;

    /**
     * 常规验证
     *
     * @param allowNull
     */
    public DateTimeVerify(String pattern, boolean allowNull) {
        super(allowNull);
        this.pattern = pattern;
    }

    /**
     * 自定义验证
     *
     * @param allowNull
     * @param customVerify
     */
    public DateTimeVerify(String pattern, boolean allowNull, Function<Object, Date> customVerify) {
        super(allowNull, customVerify);
        this.pattern = pattern;
    }

    @Override
    public Date doVerify(String fieldName, Object cellValue) throws Exception {
        if (cellValue instanceof Date) {
            return (Date) cellValue;
        }
        SimpleDateFormat format = new SimpleDateFormat(pattern);
        return format.parse(String.valueOf(cellValue));

    }

}
