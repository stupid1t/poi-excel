package com.github.stupdit1t.excel.common;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.function.Supplier;

/**
 * 工作簿类型
 */
public enum PoiWorkbookType {

    /**
     * 大数据工作簿
     * <p>
     * 速度慢, 可以解决VM内存不够用问题, 单sheet最大1048576行
     */
    BIG_XLSX(XSSFWorkbook::new),

    /**
     * 07 EXCEL
     * <p>
     * 速度慢, 单sheet最大1048576行
     */
    XLSX(XSSFWorkbook::new),

    /**
     * 03 Excel
     * <p>
     * 速度较快, 单sheet最大65535行
     */
    XLS(HSSFWorkbook::new),
    ;

    /**
     * the number of rows that are kept in memory until flushed out, see above.
     */
    private int rowAccessWindowSize = 200;

    /**
     * whether to use gzip compression for temporary files
     */
    private boolean compressTmpFiles;

    /**
     * whether to use a shared strings table
     */
    private boolean useSharedStringsTable;

    /**
     * 创建工作簿方法
     */
    private final Supplier<Workbook> create;

    PoiWorkbookType(Supplier<Workbook> create) {
        this.create = create;
    }

    /**
     * 创建工作簿
     *
     * @return Workbook
     */
    public Workbook create() {
        Workbook workbook = create.get();
        if (this == PoiWorkbookType.BIG_XLSX) {
            workbook = new SXSSFWorkbook((XSSFWorkbook) workbook, this.rowAccessWindowSize, this.compressTmpFiles, this.useSharedStringsTable);
        }
        return workbook;
    }

    /**
     * 设置内存行数
     *
     * @param rowAccessWindowSize 默认100
     * @return PoiWorkbookType
     */
    public PoiWorkbookType rowAccessWindowSize(int rowAccessWindowSize) {
        this.rowAccessWindowSize = rowAccessWindowSize;
        return this;
    }

    /**
     * 压缩临时文件
     *
     * @param compressTmpFiles 默认false
     * @return PoiWorkbookType
     */
    public PoiWorkbookType compressTmpFiles(boolean compressTmpFiles) {
        this.compressTmpFiles = compressTmpFiles;
        return this;
    }

    /**
     * 使用共享字符串表
     *
     * @param useSharedStringsTable 默认false
     * @return PoiWorkbookType
     */
    public PoiWorkbookType useSharedStringsTable(boolean useSharedStringsTable) {
        this.useSharedStringsTable = useSharedStringsTable;
        return this;
    }
}
