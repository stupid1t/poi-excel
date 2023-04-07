package com.github.stupdit1t.excel.core.parse;

import com.github.stupdit1t.excel.callback.CellCallback;
import com.github.stupdit1t.excel.core.AbsParent;
import com.github.stupdit1t.excel.handle.*;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;

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
     * 验证规则
     */
    BaseVerifyRule<?, R> cellVerifyRule = new StringHandler<R>(true, this.parent);

    public InColumn(OpsColumn<R> opsColumn, String index, String field) {
        super(opsColumn);
        this.index = index;
        this.field = field;
    }

    /**
     * 必须为字符串
     *
     * @return InColumn<R>
     */
    public StringHandler<R> asString() {
        this.cellVerifyRule = new StringHandler<>(true, this.parent);
        return (StringHandler<R>) this.cellVerifyRule;
    }

    /**
     * 必须为int类型
     *
     * @return InColumn<R>
     */
    public IntegerHandler<R> asInt() {
        this.cellVerifyRule = new IntegerHandler<>(true, this.parent);
        return (IntegerHandler<R>) this.cellVerifyRule;
    }

    /**
     * 必须为long类型
     *
     * @return InColumn<R>
     */
    public LongHandler<R> asLong() {
        this.cellVerifyRule = new LongHandler<>(true, this.parent);
        return (LongHandler<R>) this.cellVerifyRule;
    }

    /**
     * 必须为布尔类型
     *
     * @return InColumn<R>
     */
    public BooleanHandler<R> asBoolean() {
        this.cellVerifyRule = new BooleanHandler<>(true, this.parent);
        return (BooleanHandler<R>) this.cellVerifyRule;
    }

    /**
     * 必须为BigDecimal类型
     *
     * @return InColumn<R>
     */
    public BigDecimalHandler<R> asBigDecimal() {
        this.cellVerifyRule = new BigDecimalHandler<>(true, this.parent);
        return (BigDecimalHandler<R>) this.cellVerifyRule;
    }

    /**
     * 必须为Char类型
     *
     * @return InColumn<R>
     */
    public CharHandler<R> asChar() {
        this.cellVerifyRule = new CharHandler<>(true, this.parent);
        return (CharHandler<R>) this.cellVerifyRule;
    }

    /**
     * 必须为Date类型
     *
     * @return InColumn<R>
     */
    public DateHandler<R> asDate() {
        this.cellVerifyRule = new DateHandler<>(true, this.parent);
        return (DateHandler<R>) this.cellVerifyRule;
    }

    /**
     * 必须为Double类型
     *
     * @return InColumn<R>
     */
    public DoubleHandler<R> asDouble() {
        this.cellVerifyRule = new DoubleHandler<>(true, this.parent);
        return (DoubleHandler<R>) this.cellVerifyRule;
    }

    /**
     * 必须为Float类型
     *
     * @return InColumn<R>
     */
    public FloatHandler<R> asFloat() {
        this.cellVerifyRule = new FloatHandler<>(true, this.parent);
        return (FloatHandler<R>) this.cellVerifyRule;
    }

    /**
     * 必须为图片类型
     *
     * @return InColumn<R>
     */
    public ImgHandler<R> asImg() {
        this.cellVerifyRule = new ImgHandler<>(true, this.parent);
        return (ImgHandler<R>) this.cellVerifyRule;
    }

    /**
     * 必须为Short类型
     *
     * @return InColumn<R>
     */
    public ShortHandler<R> asShort() {
        this.cellVerifyRule = new ShortHandler<>(true, this.parent);
        return (ShortHandler<R>) this.cellVerifyRule;
    }

    /**
     * 自定义处理
     *
     * @param handle row 行号
     *               col 列号
     *               value 当前字段值
     */
    public ObjectHandler<R> asByCustom(CellCallback handle) {
        this.cellVerifyRule = new ObjectHandler(true, this.parent, handle);
        return (ObjectHandler<R>) this.cellVerifyRule;
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
     * 获取校验规则
     *
     * @return BaseVerifyRule
     */
    public BaseVerifyRule<?, R> getCellVerifyRule() {
        return cellVerifyRule;
    }
}
