package com.github.stupdit1t.excel.common;

import org.apache.commons.lang3.StringUtils;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 常量定义
 *
 * @author 625
 */
public class POIConstant {

    /**
     * 单元格坐标对应数字Map
     */
    public static Map<String, Integer> cellRefNums = new LinkedHashMap<>();

    /**
     * 数字列对应字母
     */
    public static Map<Integer, String> numsRefCell = new HashMap<>();

    /**
     * 字母最大个数
     */
    public static final int CHAR_MAX = 26;

    static {
        // 填充2列A~ZZ
        fillCellRefNums(2);
    }

    /**
     * 列宽单位-字符
     */
    private final static int CHAR_UNIT = 2 * 310;

    /**
     * 格式化(24小时制)<br>
     * FORMAT_DateTime: 日期时间 yyyy-MM-dd HH:mm
     */
    public final static String FMT_DATE_TIME = "yyyy-MM-dd HH:mm";

    /**
     * 格式化(24小时制)<br>
     * FORMAT_DateTime: 日期时间 yyyy-MM-dd
     */
    public final static String FMT_DATE = "yyyy-MM-dd";


    /**
     * 宽度设置,
     *
     * @param charNum 汉字数量
     * @return int
     */
    public static int width(int charNum) {
        return CHAR_UNIT * charNum;
    }

    /**
     * 转换列坐标为数字
     *
     * @param cellRefs 列坐标
     * @return int[]
     */
    public static int[] convertToCellNum(String[] cellRefs) {
        int[] nums = new int[cellRefs.length];
        for (int i = 0; i < cellRefs.length; i++) {
            nums[i] = cellRefNums.get(cellRefs[i]);
        }
        return nums;
    }

    /**
     * 转换数字为坐标
     *
     * @param cellNum 列坐标
     * @return int[]
     */
    public static String convertToCellChar(Integer cellNum) {
        return numsRefCell.get(cellNum);
    }

    /**
     * 填充几位字母
     *
     * @param times
     */
    public static void fillCellRefNums(int times) {
        if (times > CHAR_MAX + 1) {
            throw new UnsupportedOperationException("最大填充27次列宽!");
        }
        String stopIndex = StringUtils.repeat('Z', times);
        String index = "A";
        int columnIndex = 0;
        cellRefNums.put(index, columnIndex);
        numsRefCell.put(columnIndex, index);
        while (true && !index.equals(stopIndex)) {
            columnIndex++;
            index = Common.charAdd(index);
            cellRefNums.put(index, columnIndex);
            numsRefCell.put(columnIndex, index);
        }
    }

}
