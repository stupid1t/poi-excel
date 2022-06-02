package com.github.stupdit1t.excel.core.parse;

import com.github.stupdit1t.excel.common.PoiCommon;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.Map;

/**
 * 导出规则定义
 */
public class OpsParse<R> {

    /**
     * 数据类型
     */
    Class<R> rowClass;

    /**
     * 实体类包含的字段
     */
    boolean mapData;

    /**
     * 实体类包含的字段
     */
    Map<String, Field> allFields;

    /**
     * 文件源类型
     * <p>
     * 1. path
     * 2. stream
     */
    int fromMode;

    /**
     * 导入来源
     */
    String fromPath;

    /**
     * 导入来源
     */
    InputStream fromStream;

    public OpsParse(Class<R> rowClass) {
        this.rowClass = rowClass;
        this.mapData = PoiCommon.isMapData(this.rowClass);
        if (!this.mapData) {
            this.allFields = PoiCommon.getAllFields(this.rowClass);
        }
    }

    /**
     * 检测是否已经被设置状态
     */
    private void checkSetFromMode(int wantSetMode) {
        if (fromMode != 0 && fromMode != wantSetMode) {
            throw new UnsupportedOperationException("仅支持设置 1 种输入方式");
        }
        this.fromMode = wantSetMode;
    }

    /**
     * 导入来源文件
     *
     * @param path 文件地址
     * @return OpsParse
     */
    public OpsParse<R> from(String path) {
        checkSetFromMode(1);
        this.fromPath = path;
        return this;
    }

    /**
     * 导入来源文件
     *
     * @param inputStream 文件流
     * @return OpsRead
     */
    public OpsParse<R> from(InputStream inputStream) {
        checkSetFromMode(2);
        this.fromStream = inputStream;
        return this;
    }

    /**
     * 导入sheet设置
     *
     * @param sheetIndex  sheet下标
     * @param headerCount 表头行数
     * @param footerCount 尾部行数
     * @return OpsSheet
     */
    public OpsSheet<R> opsSheet(int sheetIndex, int headerCount, int footerCount) {
        return new OpsSheet<>(this, sheetIndex, headerCount, footerCount);
    }

    /**
     * 导入sheet设置
     *
     * @param sheetName   sheet名字
     * @param headerCount 表头行数
     * @param footerCount 尾部行数
     * @return
     */
    public OpsSheet<R> opsSheet(String sheetName, int headerCount, int footerCount) {
        return new OpsSheet<>(this, sheetName, headerCount, footerCount);
    }
}