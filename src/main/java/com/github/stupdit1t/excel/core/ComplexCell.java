package com.github.stupdit1t.excel.core;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import java.util.function.BiConsumer;

/**
 * 复杂表头定义
 */
class ComplexCell {

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

    public String getText() {
        return text;
    }

    public BiConsumer<Font, CellStyle> getStyle() {
        return style;
    }

    public Integer[] getLocationIndex() {
        return locationIndex;
    }
}