package com.github.stupdit1t.excel.core.export;

import org.apache.poi.ss.usermodel.Workbook;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;

/**
 * 一些终结操作
 */
public interface OpsFinish {

    /**
     * 输出路径设置
     *
     * @param toPath 输出磁盘路径
     */
    void export(String toPath);

    /**
     * 输出流
     *
     * @param toStream 输出流
     */
    void export(OutputStream toStream);

    /**
     * 输出servlet
     *
     * @param toResponse 输出servlet
     * @param fileName   文件名
     */
    void export(HttpServletResponse toResponse, String fileName);

    /**
     * 执行输出
     *
     * @param workbook 导出workbook
     */
    void export(Workbook workbook);
}
