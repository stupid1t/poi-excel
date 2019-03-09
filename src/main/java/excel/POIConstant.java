package excel;

import java.util.HashMap;
import java.util.Map;

/**
 * 
 * @author 625
 *
 */
public class POIConstant {

	// 单元格坐标对应数字Map
	public static Map<String, Integer> cellRefNums = new HashMap<>();

	static {
		cellRefNums.put("A", 0);
		cellRefNums.put("B", 1);
		cellRefNums.put("C", 2);
		cellRefNums.put("D", 3);
		cellRefNums.put("E", 4);
		cellRefNums.put("F", 5);
		cellRefNums.put("G", 6);
		cellRefNums.put("H", 7);
		cellRefNums.put("I", 8);
		cellRefNums.put("J", 9);
		cellRefNums.put("K", 10);
		cellRefNums.put("L", 11);
		cellRefNums.put("M", 12);
		cellRefNums.put("N", 13);
		cellRefNums.put("O", 14);
		cellRefNums.put("P", 15);
		cellRefNums.put("Q", 16);
		cellRefNums.put("R", 17);
		cellRefNums.put("S", 18);
		cellRefNums.put("T", 19);
		cellRefNums.put("U", 20);
		cellRefNums.put("V", 21);
		cellRefNums.put("W", 22);
		cellRefNums.put("X", 23);
		cellRefNums.put("Y", 24);
		cellRefNums.put("Z", 25);
		cellRefNums.put("AA", 26);
		cellRefNums.put("AB", 27);
		cellRefNums.put("AC", 28);
		cellRefNums.put("AD", 29);
		cellRefNums.put("AE", 30);
		cellRefNums.put("AF", 31);
		cellRefNums.put("AG", 32);
		cellRefNums.put("AH", 33);
		cellRefNums.put("AI", 34);
		cellRefNums.put("AJ", 35);
		cellRefNums.put("AK", 36);
		cellRefNums.put("AL", 37);
		cellRefNums.put("AM", 38);
		cellRefNums.put("AN", 39);
		cellRefNums.put("AO", 40);
		cellRefNums.put("AP", 41);
		cellRefNums.put("AQ", 42);
		cellRefNums.put("AR", 43);
		cellRefNums.put("AS", 44);
		cellRefNums.put("AT", 45);
		cellRefNums.put("AU", 46);
		cellRefNums.put("AV", 47);
		cellRefNums.put("AW", 48);
		cellRefNums.put("AX", 49);
		cellRefNums.put("AY", 50);
		cellRefNums.put("AZ", 51);
	}

	/**
	 * 列宽单位-字符
	 */
	private final static int CHARUNIT = 2 * 310;

	/**
	 * 格式化(24小时制)<br>
	 * FORMAT_DateTime: 日期时间 yyyy-MM-dd HH:mm
	 */
	public final static String FMTDATETIME = "yyyy-MM-dd HH:mm";

	/**
	 * 格式化(24小时制)<br>
	 * FORMAT_DateTime: 日期时间 yyyy-MM-dd
	 */
	public final static String FMTDATE = "yyyy-MM-dd";


	/**
	 * 宽度设置,
	 * 
	 * @param charNum 汉字数量
	 * @return int
	 */
	public static int width(int charNum) {
		return CHARUNIT * charNum;
	}

	/**
	 * 转换列坐标为数字
	 * 
	 * @param cellRefs 列坐标
	 * @return  int[]
	 */
	public static int[] convertToCellNum(String[] cellRefs) {
		int[] nums = new int[cellRefs.length];
		for (int i = 0; i < cellRefs.length; i++) {
			nums[i] = cellRefNums.get(cellRefs[i]);
		}
		return nums;
	}

}
