package com.github.stupdit1t.excel.core.export;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import java.util.function.BiConsumer;

/**
 * 复杂表头定义
 */
public class ComplexCell {

    /**
     * header原生坐标, 如0,0,0,0
     */
    Integer[] locationIndex;

    /**
     * 样式定义
     */
    BiConsumer<Font, CellStyle> style;

    /**
     * 显示内容
     */
    String text;

    /**
     * 获取显示内容
     *
     * @return String
     */
    public String getText() {
        return text;
    }

    /**
     * 获取样式
     *
     * @return BiConsumer
     */
    public BiConsumer<Font, CellStyle> getStyle() {
        return style;
    }

    /**
     * 获取单元格定位
     *
     * @return Integer[]
     */
    public Integer[] getLocationIndex() {
        return locationIndex;
    }
}