package com.github.stupdit1t.excel.core;

import com.github.stupdit1t.excel.style.DefaultCellStyleEnum;
import com.github.stupdit1t.excel.style.ICellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;

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
    boolean xls;

    /**
     * 全局单元格样式
     */
    ICellStyle[] style = DefaultCellStyleEnum.values();

    /**
     * Excel密码, 只支持xls 格式
     */
    String password;

    /**
     * 输出模式
     * 1 路径输出
     * 2 流输出
     */
    int outMode;

    /**
     * 输出目录
     */
    String outPath;

    OpsExport() {

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
    private void checkSetOutMode(int wantSetMode) {
        if (outMode != 0 && outMode != wantSetMode) {
            throw new UnsupportedOperationException("仅支持设置输出 1 种输出方式");
        }
        this.outMode = wantSetMode;
    }

    /**
     * 输出路径设置
     *
     * @param xls 是否为xls格式
     * @return OpsExport
     */
    public OpsExport xls(boolean xls) {
        this.xls = xls;
        return this;
    }

    /**
     * 输出路径设置
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
     * 输出路径设置
     *
     * @param outPath 输出磁盘路径
     */
    public void out(String outPath) {
        checkSetOutMode(1);
        this.outPath = outPath;
        this.execute();
    }


    /**
     * 执行输出
     */
    void execute() {
        Workbook workbook = ExcelUtil.createEmptyWorkbook(!xls);
        for (OpsSheet<?> opsSheet : opsSheets) {
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
            // footer内容提取
            if (opsSheet.opsFooter != null) {
                List<ComplexCell> footer = opsSheet.opsFooter.complexFooter;
                exportRules.setFooterRules(footer);
            }
            // 表头高度处理
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
            ExcelUtil.fillBook(workbook, opsSheet.data, exportRules);
        }


        ExcelUtil.export(workbook, this.outPath);
    }
}