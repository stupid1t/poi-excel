package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.lang3.time.DateUtils;

import java.math.BigDecimal;
import java.util.Date;
import java.util.function.Function;


/**
 * 日期校验实体
 *
 * @author 625
 */
public class DateHandler extends AbsCellVerifyRule<Date> {

    /**
     * 日期格式
     */
    private final String pattern;

    /**
     * 常规验证
     *
     * @param allowNull 可为空
     * @param pattern   格式化
     */
    public DateHandler(boolean allowNull, String pattern) {
        super(allowNull);
        this.pattern = pattern;
    }

    /**
     * 自定义验证
     *
     * @param allowNull    可为空
     * @param pattern      格式化
     * @param customVerify 自定义校验
     */
    public DateHandler(boolean allowNull, String pattern, Function<Object, Date> customVerify) {
        super(allowNull, customVerify);
        this.pattern = pattern;
    }

    @Override
    public Date doHandle(String fieldName, Object cellValue) throws Exception {
        if (cellValue instanceof Date) {
            // 如果是日期格式通过
            return (Date) cellValue;
        } else if (NumberUtils.isCreatable(String.valueOf(cellValue))) {
            // 如果是数字
            String value = String.valueOf(cellValue);
            long date = new BigDecimal(value).longValue();
            if (value.length() == 10) {
                date *= 1000;
            }
            return new Date(date);
        } else if (cellValue instanceof String) {
            // 如果是字符串
            String value = (String) cellValue;
            return DateUtils.parseDate(value, pattern);
        }
        throw PoiException.error(fieldName+"格式不正确");
    }

}