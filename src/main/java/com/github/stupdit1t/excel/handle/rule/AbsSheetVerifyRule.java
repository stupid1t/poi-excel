package com.github.stupdit1t.excel.handle.rule;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

/**
 * 校验规则
 *
 * @author 625
 */
public abstract class AbsSheetVerifyRule {

    /**
     * 字段校验集
     */
    private List<CellVerifyRule> cellVerifyRules = new ArrayList<>();

    /**
     * 字段名称
     */
    private String[] fields;

    /**
     * key:cellName, value:对应的校验类
     */
    private Map<String, CellVerifyRule> columnVerifyRule;

    /**
     * 列坐标
     */
    private String[] cellRefs;

    /**
     * 添加规则
     *
     * @param index
     * @param field
     * @param filedName
     * @param cellVerify
     */
    public void addRule(String index, String field, String filedName, AbsCellVerifyRule cellVerify) {
        CellVerifyRule cellVerifyRule = new CellVerifyRule(index, field, filedName, cellVerify);
        this.cellVerifyRules.add(cellVerifyRule);
    }

    /**
     * 初始化规则
     */
    public void init() {
        buildRule();
        // 1、初始化filedNames
        fields = new String[cellVerifyRules.size()];
        // 2、初始化cellRefs
        cellRefs = new String[cellVerifyRules.size()];
        // 3、初始化verifys
        columnVerifyRule = new HashMap<>(cellVerifyRules.size());
        for (int i = 0; i < cellVerifyRules.size(); i++) {
            CellVerifyRule item = cellVerifyRules.get(i);
            columnVerifyRule.put(item.getField(), item);
            cellRefs[i] = item.getIndex();
            fields[i] = item.getField();
        }
    }

    /**
     * 校验
     *
     * @param filed
     * @param fileValue
     * @return Object
     * @throws Exception
     */
    public Object verify(String filed, Object fileValue) throws Exception {
        if (columnVerifyRule.containsKey(filed)) {
            CellVerifyRule cellVerifyRule = columnVerifyRule.get(filed);
            if (cellVerifyRule.getCellVerify() != null) {
                return cellVerifyRule.getCellVerify().handle(cellVerifyRule.getFieldName(), fileValue);
            }
        }
        return fileValue;
    }

    public List<CellVerifyRule> getCellVerifyRules() {
        return cellVerifyRules;
    }

    public String[] getFields() {
        return fields;
    }

    public Map<String, CellVerifyRule> getColumnVerifyRule() {
        return columnVerifyRule;
    }

    public String[] getCellRefs() {
        return cellRefs;
    }

    /**
     * 构建导入规则
     */
    protected abstract void buildRule();

    /**
     * 匿名抽象类规则
     * @param absSheetVerifyRuleConsumer
     * @return AbsSheetVerifyRule
     */
    public static AbsSheetVerifyRule buildRule(Consumer<AbsSheetVerifyRule> absSheetVerifyRuleConsumer){
        AbsSheetVerifyRule absSheetVerifyRule = new AbsSheetVerifyRule() {
            @Override
            protected void buildRule() {
                absSheetVerifyRuleConsumer.accept(this);
            }
        };
        return absSheetVerifyRule;
    }
}
