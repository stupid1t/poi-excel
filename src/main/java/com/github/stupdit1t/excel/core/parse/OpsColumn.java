package com.github.stupdit1t.excel.core.parse;

import com.github.stupdit1t.excel.common.Col;
import com.github.stupdit1t.excel.common.Fn;
import com.github.stupdit1t.excel.common.PoiCommon;
import com.github.stupdit1t.excel.core.AbsParent;

import java.beans.Introspector;
import java.lang.invoke.SerializedLambda;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;
import java.util.regex.Pattern;

/**
 * 数据列定义
 *
 * @param <R>
 */
public class OpsColumn<R> extends AbsParent<OpsSheet<R>> {

    private static final Pattern GET_PATTERN = Pattern.compile("^get[A-Z].*");

    private static final Pattern IS_PATTERN = Pattern.compile("^is[A-Z].*");

    /**
     * 导入列定义
     */
    Map<String, InColumn<?>> columns = new HashMap<>();

    boolean autoField;

    OpsColumn(OpsSheet<R> export, boolean autoField) {
        super(export);
        this.autoField = autoField;
        /**
         * map 自动填充，在最终知道列数
         */
        if (this.autoField && !this.parent.parent.mapData) {
            Map<String, Field> allFields = this.parent.parent.allFields;
            Set<Map.Entry<String, Field>> entries = allFields.entrySet();
            int index = 0;
            for (Map.Entry<String, Field> entry : entries) {
                String colChar = PoiCommon.convertToCellChar(index);
                this.field(colChar, entry.getKey());
                index++;
            }
        }
    }

    /**
     * 列字段定义
     *
     * @param index 下标, 如A/B/C/D
     * @param field 对应的字段, 支持级联设置
     * @return InColumn
     */
    public IParseRule<R> field(String index, String field) {
        // 检测字段是否存在
        if (!this.parent.parent.mapData && this.parent.parent.allFields != null) {
            if (!this.parent.parent.allFields.containsKey(field)) {
                throw new UnsupportedOperationException("字段不存在!");
            }
        }
        InColumn<R> inColumn = new InColumn<>(this, index, field);
        columns.put(index, inColumn);
        return inColumn.getCellVerifyRule();
    }

    /**
     * 列字段定义
     *
     * @param index    下标, 如A/B/C/D
     * @param fieldFun 对应的字段
     * @return InColumn
     */
    public <F> IParseRule<R> field(String index, Fn<R, F> fieldFun) {
        try {
            String field = getField(fieldFun);
            return this.field(index, field);
        } catch (ReflectiveOperationException var4) {
            throw new UnsupportedOperationException("field 字段设置异常", var4);
        }
    }

    /**
     * 列字段定义
     *
     * @param index 下标, 如A/B/C/D
     * @param field 对应的字段, 支持级联设置
     * @return InColumn
     */
    public IParseRule<R> field(Col index, String field) {
        return this.field(index.name(), field);
    }

    /**
     * 列字段定义
     *
     * @param index    下标, 如A/B/C/D
     * @param fieldFun 对应的字段
     * @return InColumn
     */
    public <F> IParseRule<R> field(Col index, Fn<R, F> fieldFun) {
        return this.field(index.name(), fieldFun);
    }

    /**
     * 获取字段
     *
     * @param fieldFun
     * @return
     * @throws NoSuchMethodException
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     */
    private String getField(Fn<R,?> fieldFun) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        Method method = fieldFun.getClass().getDeclaredMethod("writeReplace");
        method.setAccessible(Boolean.TRUE);
        SerializedLambda serializedLambda = (SerializedLambda) method.invoke(fieldFun);
        String getter = serializedLambda.getImplMethodName();
        if (GET_PATTERN.matcher(getter).matches()) {
            getter = getter.substring(3);
        } else if (IS_PATTERN.matcher(getter).matches()) {
            getter = getter.substring(2);
        }
        return Introspector.decapitalize(getter);
    }
}
