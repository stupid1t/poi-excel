package com.github.stupdit1t.excel.core;

import com.github.stupdit1t.excel.common.PoiWorkbookType;
import com.github.stupdit1t.excel.core.export.OpsExport;
import com.github.stupdit1t.excel.core.parse.OpsParse;
import com.github.stupdit1t.excel.core.replace.OpsReplace;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 快速构建导出导入类
 */
public final class ExcelHelper {

    private ExcelHelper() {

    }

    /**
     * 导出入口
     *
     * @return OpsExport
     */
    public static OpsExport opsExport(PoiWorkbookType workbookType) {
        return new OpsExport(workbookType);
    }

    /**
     * 导出入口
     *
     * @return OpsExport
     */
    public static OpsExport opsExport(Workbook workbook) {
        return new OpsExport(workbook);
    }

    /**
     * 导入入口
     *
     * @return OpsExport
     */
    public static <R> OpsParse<R> opsParse(Class<R> rowClass) {
        return new OpsParse<>(rowClass);
    }

    /**
     * 读模板替换变量入口
     *
     * @return OpsReplace
     */
    public static OpsReplace opsReplace() {
        return new OpsReplace();
    }
}
