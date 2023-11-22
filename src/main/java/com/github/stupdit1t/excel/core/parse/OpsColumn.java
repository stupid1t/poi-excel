package com.github.stupdit1t.excel.core.parse;

import com.github.stupdit1t.excel.common.Col;
import com.github.stupdit1t.excel.common.Fn;
import com.github.stupdit1t.excel.common.PoiCommon;
import com.github.stupdit1t.excel.core.AbsParent;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * 数据列定义
 *
 * @param <R>
 */
public class OpsColumn<R> extends AbsParent<OpsSheet<R>> {

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
        String field = PoiCommon.getField(fieldFun);
        return this.field(index, field);
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
    public IParseRule<R> field(Col index, Fn<R, ?> fieldFun) {
        return this.field(index.name(), fieldFun);
    }
}
