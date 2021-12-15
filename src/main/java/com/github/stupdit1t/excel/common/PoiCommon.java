package com.github.stupdit1t.excel.common;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 一些公用的方法
 */
public class PoiCommon {

    /**
     * 合并单元格转换
     *
     * @param rule 合并规则 如将1,1,A,A 转换
     * @return Object[]
     */
    public static Object[] coverRange(String rule) {
        String[] range = rule.split(",");
        Object[] rangeInt = new Object[4];
        for (int i = 0; i < range.length; i++) {
            if (i < 2) {
                rangeInt[i] = Integer.parseInt(range[i]);
            } else {
                rangeInt[i] = range[i];
            }

        }
        return rangeInt;
    }

    /**
     * 获取实体的所有字段
     *
     * @param t
     * @return Map<String, Field>
     */
    public static Map<String, Field> getAllFields(Class<?> t) {
        Map<String, Field> field = new HashMap<>();
        List<Field> allFieldsList = FieldUtils.getAllFieldsList(t);
        allFieldsList.stream().forEach(n -> {
            n.setAccessible(true);
            field.put(n.getName(), n);
        });
        return field;
    }

    /**
     * 字母相加
     *
     * @param charStr
     */
    public static String charAdd(String charStr) {
        char[] chars = charStr.toCharArray();
        for (int i = chars.length - 1; i >= 0; i--) {
            if (i != 0 && chars[i] == 90) {
                chars[i] = 'A';
                continue;
            }
            // 位数满了, 需要增加位数
            if (i == 0 && chars[i] == 90) {
                chars = new char[chars.length + 1];
                for (int j = 0; j < chars.length; j++) {
                    chars[j] = 'A';
                }
                break;
            }
            chars[i] = (char) (chars[i] + 1);
            break;
        }
        return new String(chars);
    }

    /**
     * 宽度设置,
     *
     * @param charNum 汉字数量
     * @return int
     */
    public static int width(int charNum) {
        return PoiConstant.CHAR_UNIT * charNum;
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
            nums[i] = PoiConstant.cellRefNums.get(cellRefs[i]);
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
        return PoiConstant.numsRefCell.get(cellNum);
    }

    /**
     * 填充几位字母的列, 如3 就是填充 A~ZZZ
     *
     * @param times
     */
    public static void fillCellRefNums(int times) {
        if (times > PoiConstant.CHAR_MAX + 1) {
            throw new UnsupportedOperationException("最大填充27次列宽!");
        }
        Map<String, Integer> cellRefNums = PoiConstant.cellRefNums;
        Map<Integer, String> numsRefCell = PoiConstant.numsRefCell;
        cellRefNums.clear();
        numsRefCell.clear();
        String stopIndex = StringUtils.repeat('Z', times);
        String index = "A";
        int columnIndex = 0;
        cellRefNums.put(index, columnIndex);
        numsRefCell.put(columnIndex, index);
        while (true && !index.equals(stopIndex)) {
            columnIndex++;
            index = PoiCommon.charAdd(index);
            cellRefNums.put(index, columnIndex);
            numsRefCell.put(columnIndex, index);
        }
    }
}
