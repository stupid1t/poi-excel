package com.github.stupdit1t.excel.core.export;

import com.github.stupdit1t.excel.common.PoiWorkbookType;
import com.github.stupdit1t.excel.core.ExcelUtil;
import com.github.stupdit1t.excel.style.DefaultCellStyleEnum;
import com.github.stupdit1t.excel.style.ICellStyle;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CountDownLatch;

/**
 * 导出规则定义
 */
public class OpsExport {

    /**
     * 输出的sheet
     */
    List<OpsSheet<?>> opsSheets;

    /**
     * 文件格式
     */
    PoiWorkbookType workbookType;

    /**
     * 全局单元格样式
     */
    ICellStyle[] style = DefaultCellStyleEnum.values();

    /**
     * Excel密码, 只支持xls 格式
     */
    String password;

    /**
     * 并行导出sheet
     */
    boolean parallelSheet = false;

    /**
     * 输出模式
     * 1 路径输出
     * 2 流输出
     * 3 servlet 响应
     */
    int toMode;

    /**
     * 输出目录
     */
    String path;

    /**
     * 输出流
     */
    OutputStream stream;

    /**
     * 输出 Servlet 响应
     */
    HttpServletResponse response;

    /**
     * 输出 Servlet 响应 文件名
     */
    String responseName;

    public OpsExport(PoiWorkbookType workbookType) {
        this.workbookType = workbookType;
    }

    /**
     * 数据设置
     *
     * @param data 数据
     * @return OpsSheet<R>
     */
    public <R> OpsSheet<R> opsSheet(List<R> data) {
        if (opsSheets == null) {
            opsSheets = new ArrayList<>();
        }
        OpsSheet<R> opsSheet = new OpsSheet<>(this);
        opsSheets.add(opsSheet);
        opsSheet.data = data;
        return opsSheet;
    }

    /**
     * 检测是否已经被设置状态
     */
    private void checkSetToMode(int wantSetMode) {
        if (toMode != 0 && toMode != wantSetMode) {
            throw new UnsupportedOperationException("仅支持设置输出 1 种输出方式");
        }
        this.toMode = wantSetMode;
    }

    /**
     * 全局样式设置
     *
     * @param styles 样式
     * @return OpsExport
     */
    public OpsExport style(ICellStyle... styles) {
        this.style = styles;
        return this;
    }

    /**
     * 设置密码
     *
     * @param password 密码
     * @return OpsExport
     */
    public OpsExport password(String password) {
        this.password = password;
        return this;
    }

    /**
     * 并行导出sheet
     *
     * @param parallelSheet
     * @return OpsExport
     */
    public OpsExport parallelSheet(boolean parallelSheet) {
        this.parallelSheet = parallelSheet;
        return this;
    }

    /**
     * 输出路径设置
     *
     * @param toPath 输出磁盘路径
     */
    public void export(String toPath) {
        checkSetToMode(1);
        this.path = toPath;
        this.export();
    }

    /**
     * 输出流
     *
     * @param toStream 输出流
     */
    public void export(OutputStream toStream) {
        checkSetToMode(2);
        this.stream = toStream;
        this.export();
    }

    /**
     * 输出servlet
     *
     * @param toResponse 输出servlet
     * @param fileName   文件名
     */
    public void export(HttpServletResponse toResponse, String fileName) {
        checkSetToMode(3);
        this.response = toResponse;
        this.responseName = fileName;
        this.export();
    }

    /**
     * 执行输出
     */
    void export() {
        // 1.声明工作簿
        Workbook workbook = workbookType.create();

        // 密码设置
        if (StringUtils.isNotBlank(this.password)) {
            ExcelUtil.encryptWorkbook(workbook, this.password);
        }

        // 2.多sheet获取
        if (this.parallelSheet) {
            CountDownLatch count = new CountDownLatch(opsSheets.size());
            opsSheets.parallelStream().forEach(opsSheet -> {
                fillBook(workbook, opsSheet);
                count.countDown();
            });
            try {
                count.await();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        } else {
            for (OpsSheet<?> opsSheet : opsSheets) {
                fillBook(workbook, opsSheet);
            }
        }

        // 5.执行导出
        switch (this.toMode) {
            case 1:
                ExcelUtil.export(workbook, this.path);
                break;
            case 2:
                ExcelUtil.export(workbook, this.stream);
                break;
            case 3:
                ExcelUtil.export(workbook, this.response, this.responseName);
                break;
        }

    }

    /**
     * 填充book
     *
     * @param workbook 工作簿
     * @param opsSheet sheet
     */
    private void fillBook(Workbook workbook, OpsSheet<?> opsSheet) {
        // 3.header获取
        OpsHeader<?> opsHeader = opsSheet.opsHeader;
        ExportRules exportRules;
        if (opsSheet.opsHeader.mode == 2) {
            // 复杂表头模式
            OpsHeader.ComplexHeader<?> complex = opsHeader.complex;
            List<ComplexCell> complexHeader = complex.headers;
            exportRules = ExportRules.complexRule(opsSheet.opsColumn.columns, complexHeader);
        } else {
            // 简单表头模式
            OpsHeader.SimpleHeader<?> simple = opsHeader.simple;
            exportRules = ExportRules.simpleRule(opsSheet.opsColumn.columns, simple.headers);
            exportRules.title(simple.title);
        }
        exportRules.titleHeight = opsSheet.titleHeight;
        exportRules.headerHeight = opsSheet.headerHeight;
        exportRules.cellHeight = opsSheet.cellHeight;
        exportRules.footerHeight = opsSheet.footerHeight;
        exportRules.sheetName = opsSheet.sheetName;
        exportRules.freezeHeader = opsSheet.opsHeader.freeze;
        exportRules.password = this.password;
        exportRules.globalStyle = this.style;
        exportRules.autoNumColumnWidth = opsSheet.autoNumColumnWidth;
        exportRules.setAutoNum(opsSheet.autoNum);

        // footer内容提取
        if (opsSheet.opsFooter != null) {
            List<ComplexCell> footer = opsSheet.opsFooter.complexFooter;
            exportRules.setFooterRules(footer);
        }

        // 4.填充表格
        ExcelUtil.fillBook(workbook, opsSheet.data, exportRules);
    }

}