package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.commons.lang3.time.DateUtils;

import java.math.BigDecimal;
import java.util.Date;


/**
 * 日期校验实体
 *
 * @author 625
 */
public class DateHandler extends BaseVerifyRule<Date> {

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

    @Override
    public Date doHandle(String fieldName, String index, Object cellValue) throws Exception {
        if (cellValue instanceof Date) {
            // 如果是日期格式通过
            Date date = (Date) cellValue;
            return StringUtils.isBlank(pattern) ? date : DateUtils.parseDate(DateFormatUtils.format(date, pattern), pattern);
        } else if (NumberUtils.isCreatable(String.valueOf(cellValue))) {
            // 如果是数字
            String value = String.valueOf(cellValue);
            long date = new BigDecimal(value).longValue();
            if (value.length() == 10) {
                date *= 1000;
            }
            if (date > 1000000000000L) {
                Date dateVal = new Date(date);
                return StringUtils.isBlank(pattern) ? dateVal : DateUtils.parseDate(DateFormatUtils.format(dateVal, pattern), pattern);
            }
        } else if (cellValue instanceof String) {
            // 如果是字符串
            String value = (String) cellValue;
            return DateUtils.parseDate(value, pattern);
        }
        throw PoiException.error(String.format(PoiConstant.INCORRECT_FORMAT_STR, fieldName, index));
    }
}