package com.github.stupdit1t.excel.handle;

import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.core.parse.OpsColumn;
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
public class DateHandler<R> extends BaseVerifyRule<Date, R> {

    /**
     * 日期格式
     */
	private String pattern;

    /**
     * 是否为1904date
     */
	private boolean is1904Date;

    /**
     * 常规验证
     *
     * @param allowNull 可为空
	 * @param opsColumn 格式化
     */
	public DateHandler(boolean allowNull, OpsColumn<R> opsColumn) {
		super(allowNull, opsColumn);
	}

	/**
	 * 格式化
	 *
	 * @param pattern 格式
	 */
	public DateHandler<R> pattern(String pattern) {
        this.pattern = pattern;
		return this;
	}

	/**
	 * 格式化
	 *
	 * @param is1904Date 格式
	 */
	public DateHandler<R> is1904Date(boolean is1904Date) {
        this.is1904Date = is1904Date;
		return this;
    }

    @Override
    protected Date doHandle(int row, int col, Object cellValue) throws Exception {
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
                return StringUtils.isBlank(pattern) ? DateUtils.parseDate(value, PoiConstant.FMT_DATE_TIME): DateUtils.parseDate(value, pattern);
            }
        }
    }
}