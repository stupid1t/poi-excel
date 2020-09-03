package com.github.stupdit1t.excel.common;

import java.util.HashMap;
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
	public static Map<String, Integer> cellRefNums = new HashMap<>();

	/**
	 * 字母最大个数
	 */
	public static final int CHAR_MAX = 26;

	static {
		fillCellRefNums(5);
	}

	/**
	 * 填充列宽，可以人为扩充
	 *
	 * @param times
	 */
	public static void fillCellRefNums(int times) {
		if (times > CHAR_MAX + 1) {
			throw new UnsupportedOperationException("最大填充27次列宽!");
		}
		int columnIndex = 0;
		String key = null;
		for (int i = 0; i < times; i++) {
			for (int j = 0; j < CHAR_MAX; j++, columnIndex++) {
				int temp = columnIndex / CHAR_MAX - 1;
				if (temp == -1) {
					key = Character.toString((char) ('A' + j));
				} else {
					key = Character.toString((char) ('A' + temp)) + (char) ('A' + j);
				}
				cellRefNums.put(key, columnIndex);
			}
		}
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

}
