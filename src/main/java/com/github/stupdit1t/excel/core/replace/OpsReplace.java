package com.github.stupdit1t.excel.core.replace;

import com.github.stupdit1t.excel.core.OpsPoiUtil;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

public class OpsReplace {

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

    /**
     * 设置密码
     */
    String password;

    /**
     * 要替换的变量
     */
    Map<String, Object> variable = new HashMap<>();

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
    public OpsReplace from(String path) {
        checkSetFromMode(1);
        this.fromPath = path;
        return this;
    }

    /**
     * 导入来源文件
     *
     * @param inputStream 文件流
     * @return OpsReplace
     */
    public OpsReplace from(InputStream inputStream) {
        checkSetFromMode(2);
        this.fromStream = inputStream;
        return this;
    }

    /**
     * 设置密码
     *
     * @param password 密码
     * @return OpsReplace
     */
    public OpsReplace password(String password) {
        this.password = password;
        return this;
    }

    /**
     * 替换变量
     *
     * @param variable 变量
     * @return OpsReplace
     */
    public OpsReplace var(Map<String, Object> variable) {
        this.variable.putAll(variable);
        return this;
    }

    /**
     * 替换变量
     *
     * @param key 变量名
     * @param value 变量值
     * @return OpsReplace
     */
    public OpsReplace var(String key, Object value) {
        this.variable.put(key, value);
        return this;
    }

    /**
     * 替换
     *
     * @return Workbook
     */
    public Workbook replace() {
        if (StringUtils.isBlank(fromPath) && fromStream == null) {
            throw new UnsupportedOperationException("请设置输入!");
        }
        final Workbook workbook;
        if (this.fromMode == 1) {
            if(this.password != null){
                workbook = OpsPoiUtil.readExcelWrite(fromPath, this.password, variable);
            }else{
                workbook = OpsPoiUtil.readExcelWrite(fromPath, variable);
            }
        } else {
            if(this.password != null){
                workbook = OpsPoiUtil.readExcelWrite(fromStream, this.password, variable);
            }else{
                workbook = OpsPoiUtil.readExcelWrite(fromStream, variable);
            }
        }
        return workbook;
    }

    /**
     * 替换并输出
     *
     * @param path 路径
     */
    public void replaceTo(String path) {
        Workbook workbook = replace();
        OpsPoiUtil.export(workbook, path, this.password);
    }

    /**
     * 替换并输出
     *
     * @param out 流
     */
    public void replaceTo(OutputStream out) {
        Workbook workbook = replace();
        OpsPoiUtil.export(workbook, out, this.password);
    }

    /**
     * 替换并输出
     *
     * @param response 响应
     * @param filename 文件名
     */
    public void replaceTo(HttpServletResponse response, String filename) {
        Workbook workbook = replace();
        OpsPoiUtil.export(workbook, response, filename, this.password);
    }
}
