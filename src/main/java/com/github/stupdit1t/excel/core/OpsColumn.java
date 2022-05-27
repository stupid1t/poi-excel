package com.github.stupdit1t.excel.core;

import java.util.ArrayList;
import java.util.List;

/**
 * 数据列定义
 *
 * @param <R>
 */
public class OpsColumn<R> extends AbsParent<OpsSheet<R>> {

    List<Column<?>> columns;

    OpsColumn(OpsSheet<R> export) {
        super(export);
    }

    public Column<R> field(String field) {
        if (columns == null) {
            columns = new ArrayList<>();
        }
        Column<R> column = new Column<>(this, field);
        columns.add(column);
        return column;
    }

    public OpsColumn<R> fields(String... fields) {
        if (columns == null) {
            columns = new ArrayList<>();
        }
        for (String field : fields) {
            Column<R> column = new Column<>(this, field);
            columns.add(column);
        }
        return this;
    }
}
