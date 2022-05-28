package com.github.stupdit1t.excel.common;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.function.Supplier;

/**
 * 工作簿类型
 */
public enum WorkbookType {

    /**
     * 大数据工作簿
     */
    BIG_XLSX(XSSFWorkbook::new),

    XLSX(XSSFWorkbook::new),

    XLS(HSSFWorkbook::new),
    ;
    private int rowAccessWindowSize = 200;
    private boolean compressTmpFiles;
    private boolean useSharedStringsTable;
    private Supplier<Workbook> create;

    WorkbookType(Supplier<Workbook> create) {
        this.create = create;
    }

    /**
     * 创建工作簿
     *
     * @return
     */
    public Workbook create() {
        Workbook workbook = create.get();
        if (this == WorkbookType.BIG_XLSX) {
            workbook = new SXSSFWorkbook((XSSFWorkbook) workbook, this.rowAccessWindowSize, this.compressTmpFiles, this.useSharedStringsTable);
        }
        return workbook;
    }

    /**
     * 设置内存行数
     *
     * @param rowAccessWindowSize 默认100
     * @return
     */
    public WorkbookType rowAccessWindowSize(int rowAccessWindowSize) {
        this.rowAccessWindowSize = rowAccessWindowSize;
        return this;
    }

    /**
     * 压缩临时文件
     *
     * @param compressTmpFiles 默认false
     * @return
     */
    public WorkbookType compressTmpFiles(boolean compressTmpFiles) {
        this.compressTmpFiles = compressTmpFiles;
        return this;
    }

    /**
     * 使用共享字符串表
     *
     * @param useSharedStringsTable 默认false
     * @return
     */
    public WorkbookType useSharedStringsTable(boolean useSharedStringsTable) {
        this.useSharedStringsTable = useSharedStringsTable;
        return this;
    }
}
