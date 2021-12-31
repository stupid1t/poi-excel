package com.github.stupdit1t.excel.handle.rule;

/**
 * 列校验实体
 *
 * @author 625
 */
public class CellVerifyRule {

    /**
     * 列坐标
     */
    private final String index;

    /**
     * 列名
     */
    private final String field;

    /**
     * 列名称
     */
    private final String fieldName;

    /**
     * 列校验
     */
    private final AbsCellVerifyRule cellVerify;

    /**
     * 构建列校验
     *
     * @param index      列坐标
     * @param field      字段
     * @param filedName  字段描述
     * @param cellVerify 验证器
     */
    public CellVerifyRule(String index, String field, String filedName, AbsCellVerifyRule cellVerify) {
        super();
        this.index = index;
        this.field = field;
        this.fieldName = filedName;
        this.cellVerify = cellVerify;
    }

    public String getIndex() {
        return index;
    }

    public String getField() {
        return field;
    }

    public String getFieldName() {
        return fieldName;
    }

    public AbsCellVerifyRule getCellVerify() {
        return cellVerify;
    }
}
