package com.github.stupdit1t.excel.common;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.util.Arrays;
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
     * @param location 合并规则 如将1,1,A,A 或者 A1:3C  转换
     * @return Integer[]
     */
    public static Integer[] coverRangeIndex(String location) {
        Integer[] rangeInt = new Integer[4];
        boolean isExcelIndex = location.contains(":");
        if (isExcelIndex) {
            CellRangeAddress address = CellRangeAddress.valueOf(location);
            rangeInt = new Integer[]{address.getFirstRow(), address.getLastRow(), address.getFirstColumn(), address.getLastColumn()};
        } else {
            String[] range = location.split(",");
            for (int i = 0; i < range.length; i++) {
                if (i < 2) {
                    rangeInt[i] = Integer.parseInt(range[i]) - 1;
                } else {
                    rangeInt[i] = PoiConstant.cellRefNums.get(range[i]);
                }
            }
        }
        return rangeInt;
    }

    /**
     * 获取实体的所有字段
     *
     * @param t 类
     * @return Map<String, Field>
     */
    public static Map<String, Field> getAllFields(Class<?> t) {
        Map<String, Field> field = new HashMap<>();
        List<Field> allFieldsList = FieldUtils.getAllFieldsList(t);
        allFieldsList.forEach(n -> {
            n.setAccessible(true);
            field.put(n.getName(), n);
        });
        return field;
    }

    /**
     * 字母相加
     *
     * @param charStr 被自增的字符
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
                Arrays.fill(chars, 'A');
                break;
            }
            chars[i] = (char) (chars[i] + 1);
            break;
        }
        return new String(chars);
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
     * @param times 填充多少列
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
        while (!index.equals(stopIndex)) {
            columnIndex++;
            index = PoiCommon.charAdd(index);
            cellRefNums.put(index, columnIndex);
            numsRefCell.put(columnIndex, index);
        }
    }

    /**
     * 获取map规则最大列和行数
     *
     * @param indexLocation 规则
     * @return int[]
     */
    public static int[] getMapRowColNum(List<Integer[]> indexLocation) {
        // 解析rules，获取最大行和最大列
        int row = 0;
        int col = 0;
        for (Integer[] range : indexLocation) {
            int a = range[1] + 1;
            int b = range[3] + 1;
            row = Math.max(a, row);
            col = Math.max(b, col);
        }
        return new int[]{row, col};
    }

    /**
     * 拷贝样式
     *
     * @param newStyle 新
     * @param newFont  新
     * @param oldStyle 旧
     * @param oldFont  旧
     */
    public static void copyStyleAndFont(CellStyle newStyle, Font newFont, CellStyle oldStyle, Font oldFont) {
        newStyle.cloneStyleFrom(oldStyle);
        PoiCommon.copyFont(newFont, oldFont);
        newStyle.setFont(newFont);
    }

    /**
     * 拷贝样式
     *
     * @param newFont 新
     * @param oldFont 旧
     */
    public static void copyFont(Font newFont, Font oldFont) {
        newFont.setBold(oldFont.getBold());
        newFont.setFontHeight(oldFont.getFontHeight());
        newFont.setFontName(oldFont.getFontName());
        newFont.setFontHeightInPoints(oldFont.getFontHeightInPoints());
        newFont.setColor(oldFont.getColor());
        newFont.setCharSet(oldFont.getCharSet());
        newFont.setItalic(oldFont.getItalic());
        newFont.setStrikeout(oldFont.getStrikeout());
        newFont.setTypeOffset(oldFont.getTypeOffset());
        newFont.setUnderline(oldFont.getUnderline());
    }

    /**
     * 判断是否是Map型数据
     *
     * @param cls 数据类
     * @return boolean
     */
    public static boolean isMapData(Class<?> cls) {
        boolean mapData = false;
        if (cls == Map.class) {
            mapData = true;
        } else {
            Class<?>[] interfaces = cls.getInterfaces();
            for (Class<?> anInterface : interfaces) {
                mapData = isMapData(anInterface);
                if(mapData){
                    break;
                }
            }
        }
        return mapData;
    }
}
