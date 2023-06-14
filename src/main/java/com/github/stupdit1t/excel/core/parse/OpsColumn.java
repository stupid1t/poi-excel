package com.github.stupdit1t.excel.core.parse;

import com.github.stupdit1t.excel.core.AbsParent;

import java.util.HashMap;
import java.util.Map;

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

    /**
     * 声明导入列
     *
     * @param export 操作的sheet
     */
    OpsColumn(OpsSheet<R> export) {
        super(export);
    }


    /**
     * 列字段定义
     *
     * @param index 下标, 如A/B/C/D
     * @param field 对应的字段
     * @return InColumn
     */
    public InColumn<R> field(String index, String field) {
        // 检测字段是否存在
        if (!this.parent.parent.mapData && this.parent.parent.allFields != null) {
            if (!this.parent.parent.allFields.containsKey(field)) {
                throw new UnsupportedOperationException("字段不存在!");
            }
        }
        InColumn<R> inColumn = new InColumn<>(this, index, field);
        columns.put(index, inColumn);
        return inColumn;
    }
}
