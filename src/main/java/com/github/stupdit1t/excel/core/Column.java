package com.github.stupdit1t.excel.core;

import com.github.stupdit1t.excel.callback.OutCallback;
import com.github.stupdit1t.excel.common.PoiConstant;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * 列的定义
 *
 * @author 625
 */
public class Column<R> extends AbsParent<OpsColumn<R>> implements Cloneable {

    /**
     * 字段名称
     */
    final String field;

    /**
     * 下拉列表数据
     */
    String[] dropdown;

    /**
     * 日期校验,请填写例如2018-08-09~2019-08-09
     */
    String verifyDate;

    /**
     * 整数数字校验,请填写例如10~30
     */
    String verifyIntNum;

    /**
     * 浮点数字校验,请填写例如10.0~30.0
     */
    String verifyFloatNum;

    /**
     * 文本长度校验
     */
    String verifyText;

    /**
     * 自定义表达式校验
     */
    String verifyCustom;

    /**
     * 批注默认为空
     */
    String comment;

    /**
     * 定义规则个数
     */
    int verifyCount;

    /**
     * 列样式
     */
    Style style = new Style();

    /**
     * 输出处理
     */
    OutCallback<R> outHandle;

    public Column(OpsColumn<R> opsColumn, String field) {
        super(opsColumn);
        this.field = field;
    }

    /**
     * 高度
     *
     * @param height 不设置默认
     * @return Column<R>
     */
    public Column<R> height(int height) {

        style.height = height;
        return this;
    }

    /**
     * 宽度
     *
     * @param width 不设置默认自动
     * @return Column<R>
     */
    public Column<R> width(int width) {

        style.width = width;
        return this;
    }

    /**
     * 水平定位
     *
     * @param align ，CellStyle 取值
     * @return Column<R>
     */
    public Column<R> align(HorizontalAlignment align) {

        style.align = align;
        return this;
    }

    /**
     * 设置字体颜色
     *
     * @param color HSSFColor,XSSFColor
     * @return Column<R>
     */
    public Column<R> color(IndexedColors color) {

        style.color = color;
        return this;
    }

    /**
     * 设置背景色
     *
     * @param backColor 背景色
     * @return Column<R>
     */
    public Column<R> backColor(IndexedColors backColor) {

        style.backColor = backColor;
        return this;
    }

    /**
     * 批注添加
     *
     * @param comment 批注添加
     * @return Column<R>
     */
    public Column<R> comment(String comment) {
        this.comment = comment;
        return this;
    }

    /**
     * 设置垂直定位
     *
     * @param valign 默认居下
     * @return Column<R>
     */
    public Column<R> valign(VerticalAlignment valign) {

        style.valign = valign;
        return this;
    }

    /**
     * 日期格式化
     *
     * @param datePattern 格式化内容
     * @return Column<R>
     */
    public Column<R> datePattern(String datePattern) {

        style.datePattern = datePattern;
        return this;
    }

    /**
     * 下拉列表数据
     *
     * @param dropDown 下拉列表数据
     * @return Column<R>
     */
    public Column<R> dropdown(String[] dropDown) {
        if (++verifyCount > 1) {
            throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
        }
        this.dropdown = dropDown;
        return this;
    }

    /**
     * 日期数据校验
     *
     * @param verifyDate 表达式，请填写例如2018-08-09~2019-08-09 格式也可以 yyyy-MM-dd HH:mm:ss
     * @param msgInfo    提示消息
     * @return Column<R>
     */
    public Column<R> verifyDate(String verifyDate, String... msgInfo) {
        if (++verifyCount > 1) {
            throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
        }
        if (msgInfo.length > 0) {
            this.verifyDate = verifyDate + "@" + msgInfo[0];
        } else {
            this.verifyDate = verifyDate;
        }

        return this;
    }

    /**
     * 整数数字数据校验
     *
     * @param verifyIntNum 表达式,请填写例如10~30
     * @param msgInfo      提示消息
     * @return Column<R>
     */
    public Column<R> verifyIntNum(String verifyIntNum, String... msgInfo) {
        if (++verifyCount > 1) {
            throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
        }
        if (msgInfo.length > 0) {
            this.verifyIntNum = verifyIntNum + "@" + msgInfo[0];
        } else {
            this.verifyIntNum = verifyIntNum;
        }
        return this;
    }

    /**
     * 浮点数字数据校验
     *
     * @param verifyFloatNum 表达式,请填写例如10.0~30.0
     * @param msgInfo        提示消息
     * @return Column<R>
     */
    public Column<R> verifyFloatNum(String verifyFloatNum, String... msgInfo) {
        if (++verifyCount > 1) {
            throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
        }
        if (msgInfo.length > 0) {
            this.verifyFloatNum = verifyFloatNum + "@" + msgInfo[0];
        } else {
            this.verifyFloatNum = verifyFloatNum;
        }
        return this;
    }

    /**
     * 自定义表达式校验
     *
     * @param verifyCustom 表达式 ， 注意！！！xls格式和xlsx格式的表达式不太一样，xls从当前位置A1开始算起，xlsx从当前位置开始算起,已经修正过了
     * @param msgInfo      提示消息
     * @return Column<R>
     */
    public Column<R> verifyCustom(String verifyCustom, String... msgInfo) {
        if (++verifyCount > 1) {
            throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
        }
        if (msgInfo.length > 0) {
            this.verifyCustom = verifyCustom + "@" + msgInfo[0];
        } else {
            this.verifyCustom = verifyCustom;
        }
        return this;
    }

    /**
     * 文本长度校验
     *
     * @param verifyText 比如输入1~2
     * @param msgInfo    提示消息
     * @return Column<R>
     */
    public Column<R> verifyText(String verifyText, String... msgInfo) {
        if (++verifyCount > 1) {
            throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
        }
        if (msgInfo.length > 0) {
            this.verifyText = verifyText + "@" + msgInfo[0];
        } else {
            this.verifyText = verifyText;
        }

        return this;
    }

    /**
     * 输出出路
     *
     * @param outHandle 处理内容
     * @return Column<R>
     */
    public Column<R> outHandle(OutCallback<R> outHandle) {
        this.outHandle = outHandle;
        return this;
    }

    /**
     * 列样式
     */
    public static class Style implements Cloneable {

        /**
         * 宽度，不设置默认自动
         */
        int width = -1;

        /**
         * 高度，设置是行的高度
         */
        int height = -1;

        /**
         * 水平定位，默认居中
         */
        HorizontalAlignment align;

        /**
         * 垂直定位，默认居下
         */
        VerticalAlignment valign;

        /**
         * 字体颜色，默认黑色
         */
        IndexedColors color;

        /**
         * 背景色，默认无
         */
        IndexedColors backColor;

        /**
         * 导出日期格式
         */
        String datePattern = PoiConstant.FMT_DATE_TIME;


        public void setWidth(int width) {
            this.width = width;
        }

        public void setHeight(int height) {
            this.height = height;
        }

        public void setAlign(HorizontalAlignment align) {
            this.align = align;
        }

        public void setValign(VerticalAlignment valign) {
            this.valign = valign;
        }

        public void setColor(IndexedColors color) {
            this.color = color;
        }

        public void setBackColor(IndexedColors backColor) {
            this.backColor = backColor;
        }

        public void setDatePattern(String datePattern) {
            this.datePattern = datePattern;
        }

        /**
         * 获取样式缓存
         *
         * @return
         */
        String getStyleCacheKey() {
            if (
                    this.width == -1
                            && this.height == -1
                            && this.align == null
                            && this.valign == null
                            && this.color == null
                            && this.backColor == null
                            && this.datePattern == null
            ) {
                return null;
            }
            return "Column{" +
                    "  width=" + width +
                    ", height=" + height +
                    ", align=" + align +
                    ", valign=" + valign +
                    ", color=" + color +
                    ", backColor=" + backColor +
                    ", datePattern='" + datePattern + '\'' +
                    '}';
        }

        /**
         * clone 对象
         *
         * @return oldStyle
         */
        public static Style clone(Style oldStyle) {
            Style style = null;
            try {
                style = (Style) oldStyle.clone();
            } catch (CloneNotSupportedException e) {
                e.printStackTrace();
            }
            return style;
        }

    }
}
