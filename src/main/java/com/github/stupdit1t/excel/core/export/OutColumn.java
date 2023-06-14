package com.github.stupdit1t.excel.core.export;

import com.github.stupdit1t.excel.callback.OutCallback;
import com.github.stupdit1t.excel.core.AbsParent;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.Collection;

/**
 * 列的定义
 *
 * @author 625
 */
public class OutColumn<R> extends AbsParent<OpsColumn<R>> implements Cloneable {

    /**
     * 字段名称
     */
    final String field;

    /**
     * 下拉列表数据
     */
    String[] dropdown;

    /**
     * 纵向数据相同合并
     */
    String[] mergerRepeatFieldValue;

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

    public OutColumn(OpsColumn<R> opsColumn, String field) {
        super(opsColumn);
        this.field = field;
    }

    /**
     * 高度
     *
     * @param height 不设置默认
     * @return Column<R>
     */
    public OutColumn<R> height(int height) {

        style.height = height;
        return this;
    }

    /**
     * 宽度
     *
     * @param width 不设置默认自动
     * @return Column<R>
     */
    public OutColumn<R> width(int width) {

        style.width = width;
        return this;
    }

    /**
     * 水平定位
     *
     * @param align ，CellStyle 取值
     * @return Column<R>
     */
    public OutColumn<R> align(HorizontalAlignment align) {

        style.align = align;
        return this;
    }

    /**
     * 设置字体颜色
     *
     * @param color HSSFColor,XSSFColor
     * @return Column<R>
     */
    public OutColumn<R> color(IndexedColors color) {

        style.color = color;
        return this;
    }

    /**
     * 设置背景色
     *
     * @param backColor 背景色
     * @return Column<R>
     */
    public OutColumn<R> backColor(IndexedColors backColor) {

        style.backColor = backColor;
        return this;
    }

    /**
     * 批注添加
     *
     * @param comment 批注添加
     * @return Column<R>
     */
    public OutColumn<R> comment(String comment) {
        this.comment = comment;
        return this;
    }

    /**
     * 设置垂直定位
     *
     * @param valign 默认居下
     * @return Column<R>
     */
    public OutColumn<R> valign(VerticalAlignment valign) {

        style.valign = valign;
        return this;
    }

    /**
     * 格式化单元格内人, 参考 BuiltinFormats 类
     *
     * @param pattern 格式化内容
     * @return Column<R>
     */
    public OutColumn<R> pattern(String pattern) {

        style.pattern = pattern;
        return this;
    }

    /**
     * 下拉列表数据
     *
     * @param dropDown 下拉列表数据
     * @return Column<R>
     */
    public OutColumn<R> dropdown(String... dropDown) {
        if (++verifyCount > 1) {
            throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
        }
        this.dropdown = dropDown;
        return this;
    }

    /**
     * 下拉列表数据
     *
     * @param dropDown 下拉列表数据
     * @return Column<R>
     */
    public OutColumn<R> dropdown(Collection<String> dropDown) {
        if (++verifyCount > 1) {
            throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
        }
        this.dropdown = dropDown.toArray(new String[]{});
        return this;
    }

    /**
     * 换行显示
     *
     * @return Column<R>
     */
    public OutColumn<R> wrapText() {
        style.wrapText = true;
        return this;
    }

    /**
     * 当前行重复合并当前行
     *
     * @return OutColumn<R>
     */
    public OutColumn<R> mergerRepeat() {
        this.mergerRepeatFieldValue = new String[]{this.field};
        return this;
    }

    /**
     * 指定字段值重复合并当前行
     *
     * @return OutColumn<R>
     */
    public OutColumn<R> mergerRepeat(String... field) {
        this.mergerRepeatFieldValue = field;
        return this;
    }

    /**
     * 日期数据校验
     *
     * @param verifyDate 表达式，请填写例如2018-08-09~2019-08-09 格式也可以 yyyy-MM-dd HH:mm:ss
     * @param msgInfo    提示消息
     * @return Column<R>
     */
    public OutColumn<R> verifyDate(String verifyDate, String... msgInfo) {
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
    public OutColumn<R> verifyIntNum(String verifyIntNum, String... msgInfo) {
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
    public OutColumn<R> verifyFloatNum(String verifyFloatNum, String... msgInfo) {
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
    public OutColumn<R> verifyCustom(String verifyCustom, String... msgInfo) {
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
    public OutColumn<R> verifyText(String verifyText, String... msgInfo) {
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
     * 输出设置
     *
     * @param map value    当前单元格值
     *            row      当前行记录
     *            style    自定义单元格样式
     *            rowIndex 数据下标
     * @return Column<R>
     */
    public OutColumn<R> map(OutCallback<R> map) {
        this.outHandle = map;
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
        String pattern;

        /**
         * 换行显示
         */
        Boolean wrapText;

        /**
         * 批注
         */
        String comment;


        /**
         * 获取批注
         *
         * @return
         */
        public String getComment() {
            return comment;
        }

        /**
         * 设置批注
         *
         * @param comment
         */
        public void setComment(String comment) {
            this.comment = comment;
        }

        /**
         * 获取样式缓存
         *
         * @return String
         */
        public String getStyleCacheKey() {
            if (
                    this.width == -1
                            && this.height == -1
                            && this.align == null
                            && this.valign == null
                            && this.color == null
                            && this.backColor == null
                            && this.pattern == null
                            && this.wrapText == null
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
                    ", pattern='" + pattern + '\'' +
                    ", wrapText='" + wrapText + '\'' +
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

        @Override
        protected Object clone() throws CloneNotSupportedException {
            return super.clone();
        }

        /**
         * 获取宽度
         *
         * @return int
         */
        public int getWidth() {
            return width;
        }

        /**
         * 获取高度
         *
         * @return int
         */
        public int getHeight() {
            return height;
        }

        /**
         * 获取水平定位
         *
         * @return HorizontalAlignment
         */
        public HorizontalAlignment getAlign() {
            return align;
        }

        /**
         * 获取垂直定位
         *
         * @return VerticalAlignment
         */
        public VerticalAlignment getValign() {
            return valign;
        }

        /**
         * 获取字体颜色
         *
         * @return IndexedColors
         */
        public IndexedColors getColor() {
            return color;
        }

        /**
         * 获取背景色
         *
         * @return IndexedColors
         */
        public IndexedColors getBackColor() {
            return backColor;
        }

        /**
         * 获取日期格式化 pattern
         *
         * @return String
         */
        public String getPattern() {
            return pattern;
        }

        /**
         * 设置宽度
         *
         * @param width 宽度
         */
        public void setWidth(int width) {
            this.width = width;
        }

        /**
         * 设置高度
         *
         * @param height 宽度
         */
        public void setHeight(int height) {
            this.height = height;
        }

        /**
         * 设置水平
         *
         * @param align 水平样式
         */
        public void setAlign(HorizontalAlignment align) {
            this.align = align;
        }

        /**
         * 设置垂直
         *
         * @param valign 垂直样式
         */
        public void setValign(VerticalAlignment valign) {
            this.valign = valign;
        }

        /**
         * 设置字体颜色
         *
         * @param color 颜色
         */
        public void setColor(IndexedColors color) {
            this.color = color;
        }

        /**
         * 设置背景色
         *
         * @param backColor 背景色
         */
        public void setBackColor(IndexedColors backColor) {
            this.backColor = backColor;
        }

        /**
         * 设置日期格式化
         *
         * @param pattern 格式
         */
        public void setPattern(String pattern) {
            this.pattern = pattern;
        }

        /**
         * 获取换行显示
         */
        public Boolean getWrapText() {
            return wrapText;
        }

        /**
         * 设置换行显示
         *
         * @param wrapText 是否换行显示
         */
        public void setWrapText(Boolean wrapText) {
            this.wrapText = wrapText;
        }
    }

    /**
     * 获取导出字段
     *
     * @return String
     */
    public String getField() {
        return field;
    }

    /**
     * 获取下拉框
     *
     * @return String[]
     */
    public String[] getDropdown() {
        return dropdown;
    }

    /**
     * 获取日期校验
     *
     * @return String
     */
    public String getVerifyDate() {
        return verifyDate;
    }

    /**
     * 获取int校验
     *
     * @return String
     */
    public String getVerifyIntNum() {
        return verifyIntNum;
    }

    /**
     * 获取float校验
     *
     * @return String
     */
    public String getVerifyFloatNum() {
        return verifyFloatNum;
    }

    /**
     * 获取校验文字
     *
     * @return String
     */
    public String getVerifyText() {
        return verifyText;
    }

    /**
     * 获取自定义校验
     *
     * @return String
     */
    public String getVerifyCustom() {
        return verifyCustom;
    }

    /**
     * 获取批注
     *
     * @return String
     */
    public String getComment() {
        return comment;
    }

    /**
     * 获取校验数量
     *
     * @return int
     */
    public int getVerifyCount() {
        return verifyCount;
    }

    /**
     * 获取样式
     *
     * @return Style
     */
    public Style getStyle() {
        return style;
    }

    /**
     * 获取输出回调
     *
     * @return OutCallback
     */
    public OutCallback<R> getOutHandle() {
        return outHandle;
    }

    /**
     * 行重复合并
     */
    public String[] getMergerRepeatFieldValue() {
        return mergerRepeatFieldValue;
    }

}
