package com.github.stupdit1t.excel.core.parse;

import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiException;
import com.github.stupdit1t.excel.core.AbsParent;
import com.github.stupdit1t.excel.common.TypeHandler;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.reflect.TypeUtils;

import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.util.Date;
import java.util.function.Function;

/**
 * 列校验和格式化接口
 *
 * @author 625
 */
public class BaseParseRule<R> extends AbsParent<OpsSheet<R>> implements IParseRule<R> {

    /**
     * 是否可为空
     */
    protected boolean allowNull = true;

    /**
     * 是否去空格
     */
    protected boolean trim;

    /**
     * 默认值
     */
    protected Object defaultValue;

    /**
     * 目标class。
     */
    protected Class<?> type;

    /**
     * 映射转换
     */
    private Function<Object, Object> mapping;

    /**
     * 正则校验
     */
    private String regex;

    /**
     * 格式化
     */
    private String format;

    /**
     * 精度
     */
    private int scale;

    /**
     * 操作列
     */
    private OpsColumn<R> opsColumn;

    /**
     * 构建校验规则
     */
    public BaseParseRule(OpsColumn<R> opsColumn, OpsSheet<R> parent) {
        super(parent);
        this.opsColumn = opsColumn;
    }

    /**
     * 不能为空
     *
     * @return InColumn<R>
     */
    @Override
    public IParseRule<R> notNull() {
        this.allowNull = false;
        return this;
    }

    /**
     * 去除两边空格
     *
     * @return InColumn<R>
     */
    @Override
    public IParseRule<R> trim() {
        this.trim = true;
        return this;
    }

    /**
     * 去除两边空格
     *
     * @return InColumn<R>
     */
    @Override
    public BaseParseRule<R> defaultValue(Object defaultValue) {
        this.defaultValue = defaultValue;
        return this;
    }


    /**
     * 转换or映射or判断
     */
    @Override
    public BaseParseRule<R> map(Function<Object, Object> mapping) {
        this.mapping = mapping;
        return this;
    }

    /**
     * 如果转map，进行强制转换，非必须
     *
     * @param type
     * @return
     */
    @Override
    public BaseParseRule<R> type(Class<?> type) {
        this.type = type;
        return this;
    }

    @Override
    public IParseRule<R> field(String index, String field) {
        return this.opsColumn.field(index, field);
    }

    /**
     * 正则校验
     *
     * @param regex 格式
     */
    @Override
    public BaseParseRule<R> regex(String regex) {
        this.regex = regex;
        return this;
    }

    /**
     * 格式化
     *
     * @param format 格式
     */
    @Override
    public BaseParseRule<R> format(String format) {
        this.format = format;
        return this;
    }

    @Override
    public IParseRule<R> scale(int scale) {
        this.scale = scale;
        return this;
    }

    /**
     * 判空处理
     *
     * @param value 列值
     */
    Object handleNull(Object value) throws PoiException {
        if (ObjectUtils.isEmpty(value)) {
            if (this.allowNull) {
                return null;
            } else {
                throw PoiException.error(PoiConstant.NOT_EMPTY_STR);
            }
        }
        return value;
    }

    /**
     * 校验单元格值
     *
     * @param cellValue 列值
     */
    public Object handle(int row, int col, Object cellValue, Type typeCls) throws Exception {
        // 空值处理
        cellValue = handleNull(cellValue);
        if (ObjectUtils.isEmpty(cellValue)) {
            return this.defaultValue;
        }

        // 数据映射转换，也可做判断
        if (mapping != null) {
            cellValue = mapping.apply(cellValue);
        }
        if (typeCls == null) {
            typeCls = this.type;
        }

        if (typeCls != null) {
            // 类型转换
            if (TypeUtils.equals(String.class, typeCls)) {
                cellValue = TypeHandler.stringValue(cellValue, this.trim, this.regex);
            } else if (TypeUtils.equals(Short.class, typeCls) || TypeUtils.equals(short.class, typeCls)) {
                cellValue = TypeHandler.shortValue(cellValue, this.trim, this.regex);
            } else if (TypeUtils.equals(Long.class, typeCls) || TypeUtils.equals(long.class, typeCls)) {
                cellValue = TypeHandler.longValue(cellValue, this.trim, this.regex);
            } else if (TypeUtils.equals(Integer.class, typeCls) || TypeUtils.equals(int.class, typeCls)) {
                cellValue = TypeHandler.intValue(cellValue, this.trim, this.regex);
            } else if (TypeUtils.equals(Float.class, typeCls) || TypeUtils.equals(float.class, typeCls)) {
                cellValue = TypeHandler.floatValue(cellValue, this.trim, this.regex);
            } else if (TypeUtils.equals(Double.class, typeCls) || TypeUtils.equals(double.class, typeCls)) {
                cellValue = TypeHandler.doubleValue(cellValue, this.trim, this.regex, this.scale);
            } else if (TypeUtils.equals(Date.class, typeCls)) {
                cellValue = TypeHandler.dateValue(cellValue, this.trim, this.regex, this.format, false);
            } else if (TypeUtils.equals(Boolean.class, typeCls) || TypeUtils.equals(boolean.class, typeCls)) {
                cellValue = TypeHandler.boolValue(cellValue, this.trim, this.regex);
            } else if (TypeUtils.equals(byte[].class, typeCls) || TypeUtils.equals(Byte[].class, typeCls)) {
                cellValue = TypeHandler.imgValue(cellValue, this.trim, this.regex);
            } else if (TypeUtils.equals(BigDecimal.class, typeCls)) {
                cellValue = TypeHandler.decimalValue(cellValue, this.trim, this.regex, this.scale);
            }
        }
        return cellValue;
    }

    public Class<?> getType() {
        return type;
    }
}
