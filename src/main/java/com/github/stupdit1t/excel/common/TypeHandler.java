package com.github.stupdit1t.excel.common;

import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiException;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.NumberToTextConverter;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Date;
import java.util.regex.Pattern;

public class TypeHandler {

    public static byte[] imgValue(Object cellValue, boolean trim, String regex) {
        if (cellValue instanceof byte[]) {
            return (byte[]) cellValue;
        } else {
            String value = stringValue(cellValue, trim, regex);
            return Base64.decodeBase64(value);
        }
    }

    public static BigDecimal decimalValue(Object cellValue, boolean trim, String regex, int precision) {
        if (cellValue instanceof BigDecimal) {
            return (BigDecimal) cellValue;
        } else {
            String value = stringValue(cellValue, trim, regex);
            return NumberUtils.toScaledBigDecimal(value, precision, RoundingMode.HALF_UP);
        }
    }

    public static Boolean boolValue(Object cellValue, boolean trim, String regex) {
        if (cellValue instanceof Boolean) {
            return (Boolean) cellValue;
        } else {
            String value = stringValue(cellValue, trim, regex);
            return Boolean.parseBoolean(value);
        }
    }
    public static Date dateValue(Object cellValue, boolean trim, String regex, String format, boolean is1904Date) throws Exception {
        if (cellValue instanceof Date) {
            // 如果是日期格式通过
            Date date = (Date) cellValue;
            return StringUtils.isBlank(format) ? date : DateUtils.parseDate(DateFormatUtils.format(date, format), format);
        } else {
            String value = stringValue(cellValue, trim, regex);
            if (NumberUtils.isCreatable(value)) {
                // 如果是数字
                BigDecimal sourceValue = new BigDecimal(value);
                long date = sourceValue.longValue();
                if (value.length() == 10) {
                    date *= 1000;
                }
                if (date > 1000000000000L) {
                    Date dateVal = new Date(date);
                    return StringUtils.isBlank(format) ? dateVal : DateUtils.parseDate(DateFormatUtils.format(dateVal, format), format);
                } else {
                    // 非标准时期数字
                    return DateUtil.getJavaDate(sourceValue.doubleValue(), is1904Date);
                }
            } else {
                // 如果是字符串
                return StringUtils.isBlank(format) ? DateUtils.parseDate(value, PoiConstant.FMT_DATE_TIME) : DateUtils.parseDate(value, format);
            }
        }
    }

    public static Double doubleValue(Object cellValue, boolean trim, String regex, int precision) {
        if (cellValue instanceof Double) {
            return (Double) cellValue;
        } else {
            String value = stringValue(cellValue, trim, regex);
            return NumberUtils.toScaledBigDecimal(value, precision, RoundingMode.HALF_UP).doubleValue();
        }
    }

    public static Float floatValue(Object cellValue, boolean trim, String regex) {
        if (cellValue instanceof Float) {
            return (Float) cellValue;
        } else {
            String value = stringValue(cellValue, trim, regex);
            return NumberUtils.toFloat(value);
        }
    }

    public static Integer intValue(Object cellValue, boolean trim, String regex) {
        if (cellValue instanceof Integer) {
            return (Integer) cellValue;
        } else {
            String value = stringValue(cellValue, trim, regex);
            return NumberUtils.toInt(value);
        }
    }

    public static Long longValue(Object cellValue, boolean trim, String regex) {
        if (cellValue instanceof Long) {
            return (Long) cellValue;
        } else {
            String value = stringValue(cellValue, trim, regex);
            return NumberUtils.toLong(value);
        }
    }

    public static Short shortValue(Object cellValue, boolean trim, String regex) {
        if (cellValue instanceof Short) {
            return (Short) cellValue;
        } else {
            String value = stringValue(cellValue, trim, regex);
            return NumberUtils.toShort(value);
        }
    }

    public static String stringValue(Object cellValue, boolean trim, String regex) {
        String value;
        // 处理数值 转为 string包含E科学计数的问题
        if (cellValue instanceof Number) {
            value = NumberToTextConverter.toText(((Number) cellValue).doubleValue());
        } else {
            value = String.valueOf(cellValue);
        }
        if (trim) {
            value = value.trim();
        }
        if (regex != null && !Pattern.matches(regex, value)) {
            throw PoiException.error(PoiConstant.INCORRECT_FORMAT_STR);
        }
        return value;
    }
}
