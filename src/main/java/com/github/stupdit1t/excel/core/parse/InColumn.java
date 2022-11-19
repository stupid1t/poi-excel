package com.github.stupdit1t.excel.core.parse;

import com.github.stupdit1t.excel.core.AbsParent;
import com.github.stupdit1t.excel.handle.*;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;

import java.util.function.BiFunction;

/**
 * 列的定义
 *
 * @author 625
 */
public class InColumn<R> extends AbsParent<OpsColumn<R>> {

    /**
     * 导入下标
     */
    final String index;

    /**
     * 字段
     */
    final String field;

    /**
     * 标题
     */
    final String title;

    /**
     * 验证规则
     */
    BaseVerifyRule<?> cellVerifyRule = new StringHandler(true);

    public InColumn(OpsColumn<R> opsColumn, String index, String field, String title) {
        super(opsColumn);
        this.index = index;
        this.field = field;
        this.title = title;
    }

    /**
     * 必须为字符串
     *
     * @return InColumn<R>
     */
    public InColumn<R> asString() {
        return asString(null);
    }

    /**
     * 必须为字符串
     *
     * @param pattern 正则校验单元格内容
     * @return InColumn<R>
     */
    public InColumn<R> asString(String pattern) {
        this.cellVerifyRule = new StringHandler(true, pattern);
        return this;
    }

    /**
     * 必须为int类型
     *
     * @return InColumn<R>
     */
    public InColumn<R> asInt() {
        this.cellVerifyRule = new IntegerHandler(true);
        return this;
    }

    /**
     * 必须为long类型
     *
     * @return InColumn<R>
     */
    public InColumn<R> asLong() {
        this.cellVerifyRule = new LongHandler(true);
        return this;
    }

    /**
     * 必须为布尔类型
     *
     * @return InColumn<R>
     */
    public InColumn<R> asBoolean() {
        this.cellVerifyRule = new BooleanHandler(true);
        return this;
    }

    /**
     * 必须为BigDecimal类型
     *
     * @return InColumn<R>
     */
    public InColumn<R> asBigDecimal() {
        this.cellVerifyRule = new BigDecimalHandler(true);
        return this;
    }

    /**
     * 必须为Char类型
     *
     * @return InColumn<R>
     */
    public InColumn<R> asChar() {
        this.cellVerifyRule = new CharHandler(true);
        return this;
    }

    /**
     * 必须为Date类型
     *
     * @return InColumn<R>
     */
    public InColumn<R> asDate() {
        return asDate(null, false);
    }

    /**
     * 必须为Date类型
     * @param is1904Date 非标准日期, 尝试设置false
     * @return
     */
    public InColumn<R> asDate(boolean is1904Date) {
        return asDate(null, is1904Date);
    }

    /**
     * 必须为Date类型
     *
     * @param pattern 日期格式类型
     * @return InColumn<R>
     */
    public InColumn<R> asDate(String pattern) {
        return asDate(pattern, false);
    }

    /**
     * 必须为Date类型
     *
     * @param pattern 日期格式类型
     * @param is1904Date 非标准日期, 尝试设置false
     * @return InColumn<R>
     */
    public InColumn<R> asDate(String pattern, boolean is1904Date) {
        this.cellVerifyRule = new DateHandler(true, pattern, is1904Date);
        return this;
    }

    /**
     * 必须为Double类型
     *
     * @return InColumn<R>
     */
    public InColumn<R> asDouble() {
        this.cellVerifyRule = new DoubleHandler(true);
        return this;
    }

    /**
     * 必须为Float类型
     *
     * @return InColumn<R>
     */
    public InColumn<R> asFloat() {
        this.cellVerifyRule = new FloatHandler(true);
        return this;
    }

    /**
     * 必须为图片类型
     *
     * @return InColumn<R>
     */
    public InColumn<R> asImg() {
        this.cellVerifyRule = new ImgHandler(true);
        return this;
    }

    /**
     * 必须为Short类型
     *
     * @return InColumn<R>
     */
    public InColumn<R> asShort() {
        this.cellVerifyRule = new ShortHandler(true);
        return this;
    }

    /**
     * 自定义类型
     *
     * @return InColumn<R>
     */
    public InColumn<R> asByCustom(BiFunction<String, Object, Object> handle) {
        this.cellVerifyRule = new ObjectHandler(true, handle);
        return this;
    }

    /**
     * 不能为空
     *
     * @return InColumn<R>
     */
    public InColumn<R> notNull() {
        this.cellVerifyRule.setAllowNull(false);
        return this;
    }

    /**
     * 去除两边空格
     *
     * @return InColumn<R>
     */
    public InColumn<R> trim() {
        this.cellVerifyRule.setTrim(true);
        return this;
    }

    /**
     * 获取下标
     *
     * @return String
     */
    public String getIndex() {
        return index;
    }

    /**
     * 获取字段
     *
     * @return String
     */
    public String getField() {
        return field;
    }

    /**
     * 获取标题
     *
     * @return String
     */
    public String getTitle() {
        return title;
    }

    /**
     * 获取校验规则
     *
     * @return BaseVerifyRule
     */
    public BaseVerifyRule<?> getCellVerifyRule() {
        return cellVerifyRule;
    }
}
