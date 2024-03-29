package com.github.stupdit1t.excel.common;

import org.apache.commons.lang3.StringUtils;

import java.util.HashMap;
import java.util.Map;

/**
 * 常量定义
 *
 * @author 625
 */
public class PoiConstant {

    /**
     * 单元格坐标对应数字Map
     */
    public static final Map<String, Integer> cellRefNums = new HashMap<>();

    /**
     * 数字列对应字母
     */
    public static final Map<Integer, String> numsRefCell = new HashMap<>();

    /**
     * 字母最大个数
     */
    public static final int CHAR_MAX = 26;

    /**
     * 列宽单位-字符
     */
    public static final int CHAR_UNIT = 2 * 310;

    /**
     * 格式化(24小时制)<br>
     * FORMAT_DateTime: 日期时间 yyyy-MM-dd HH:mm
     */
    public static final String FMT_DATE_TIME = "yyyy-MM-dd HH:mm:ss";

    /**
     * 格式化(24小时制)<br>
     * FORMAT_DateTime: 日期时间 yyyy-MM-dd
     */
    public static final String FMT_DATE = "yyyy-MM-dd";

    /**
     * 生成模板时, 最大填充下拉框数量
     */
    public static final int MAX_FILL_COL = 5000;

    /**
     * 不能为空模板
     */
    public static final String NOT_EMPTY_STR = "不能为空";

    /**
     * 格式不正确
     */
    public static final String INCORRECT_FORMAT_STR = "格式不正确";

    /**
     * 不能为空模板
     */
    public static final String ROW_INDEX_STR = "[%s]: %s";

    static {
        // 填充2列A~ZZ
        String times = System.getProperty("com.github.stupdit1t.fillCellTimes");
        if (StringUtils.isBlank(times)) {
            times = "2";
        }
        System.setProperty("com.github.stupdit1t.fillCellTimes", "2");
        PoiCommon.fillCellRefNums(Integer.parseInt(times));
    }
}
