package com.github.stupdit1t.excel.core.export;

import com.github.stupdit1t.excel.common.PoiCommon;
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
    private Integer[] locationIndex;

    /**
     * 样式定义
     */
    private BiConsumer<Font, CellStyle> style;

    /**
     * 显示内容
     */
    private String text;

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

    /**
     * 设置坐标
     *
     * @param locationIndex excel坐标
     */
    public void setLocationIndex(Integer[] locationIndex) {
        this.locationIndex = locationIndex;
    }

    /**
     * 设置坐标
     *
     * @param location excel坐标
     */
    public void setLocationIndex(String location) {
        this.locationIndex = PoiCommon.coverRangeIndex(location);
    }

    /**
     * 设置样式
     *
     * @param style 样式
     */
    public void setStyle(BiConsumer<Font, CellStyle> style) {
        this.style = style;
    }

    /**
     * 设置文本
     *
     * @param text 显示文本
     */
    public void setText(String text) {
        this.text = text;
    }
}