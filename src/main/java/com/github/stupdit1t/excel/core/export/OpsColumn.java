package com.github.stupdit1t.excel.core.export;

import com.github.stupdit1t.excel.core.AbsParent;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * 导出数据列定义
 *
 * @param <R>
 */
public class OpsColumn<R> extends AbsParent<OpsSheet<R>> {

    /**
     * 导出的列
     */
    List<OutColumn<?>> columns;

    /**
     * 声明
     *
     * @param export sheet
     */
    OpsColumn(OpsSheet<R> export) {
        super(export);
    }

    /**
     * 导出字段
     *
     * @param field 字段
     * @return OutColumn
     */
    public OutColumn<R> field(String field) {
        if (columns == null) {
            columns = new ArrayList<>();
        }
        OutColumn<R> column = new OutColumn<>(this, field);
        columns.add(column);
        return column;
    }

    /**
     * 字段
     *
     * @param fields 字段
     * @return OpsColumn
     */
    public OpsColumn<R> fields(String... fields) {
        if (columns == null) {
            columns = new ArrayList<>();
        }
        for (String field : fields) {
            OutColumn<R> column = new OutColumn<>(this, field);
            columns.add(column);
        }
        return this;
    }

    /**
     * 字段
     *
     * @param fields 字段
     * @return OpsColumn
     */
    public OpsColumn<R> fields(Collection<String> fields) {
        return fields(fields.toArray(new String[]{}));
    }
}
