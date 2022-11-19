package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.ss.usermodel.DateUtil;

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
     * 是否为1904date
     */
    private final boolean is1904Date;

    /**
     * 常规验证
     *
     * @param allowNull 可为空
     * @param pattern   格式化
     */
    public DateHandler(boolean allowNull, String pattern, boolean is1904Date) {
        super(allowNull);
        this.pattern = pattern;
        this.is1904Date = is1904Date;
    }

    @Override
    public Date doHandle(String fieldName, String index, Object cellValue) throws Exception {
        if (cellValue instanceof Date) {
            // 如果是日期格式通过
            Date date = (Date) cellValue;
            return StringUtils.isBlank(pattern) ? date : DateUtils.parseDate(DateFormatUtils.format(date, pattern), pattern);
        } else {
            String value = String.valueOf(cellValue);
            if (this.trim) {
                value = value.trim();
            }
            if (NumberUtils.isCreatable(value)) {
                // 如果是数字
                BigDecimal sourceValue = new BigDecimal(value);
                long date = sourceValue.longValue();
                if (value.length() == 10) {
                    date *= 1000;
                }
                if (date > 1000000000000L) {
                    Date dateVal = new Date(date);
                    return StringUtils.isBlank(pattern) ? dateVal : DateUtils.parseDate(DateFormatUtils.format(dateVal, pattern), pattern);
                } else {
                    // 非标准时期数字
                    return DateUtil.getJavaDate(sourceValue.doubleValue(), this.is1904Date);
                }
            } else {
                // 如果是字符串
                return DateUtils.parseDate(value, pattern);
            }
        }
    }
}