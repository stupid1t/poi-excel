package com.github.stupdit1t.excel.verify.rule;

/**
 * 列校验实体
 *
 * @author 625
 */
public class CellVerifyRule {

    /**
     * 列坐标
     */
    private String index;

    /**
     * 列名
     */
    private String field;

    /**
     * 列名称
     */
    private String fieldName;

    /**
     * 列校验
     */
    private AbsCellVerifyRule cellVerify;

    /**
     * 构建列校验
     *
     * @param index     列坐标
     * @param field     字段
     * @param filedName 字段名称
     */
    public CellVerifyRule(String index, String field, String filedName) {
        super();
        this.field = field;
        this.index = index;
        this.fieldName = filedName;
    }

    /**
     * 构建列校验
     *
     * @param index     列坐标
     * @param field     字段
     * @param filedName 字段名称
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

    public void setIndex(String index) {
        this.index = index;
    }

    public String getField() {
        return field;
    }

    public void setField(String field) {
        this.field = field;
    }

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public AbsCellVerifyRule getCellVerify() {
        return cellVerify;
    }

    public void setCellVerify(AbsCellVerifyRule cellVerify) {
        this.cellVerify = cellVerify;
    }
}
