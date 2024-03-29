package com.github.stupdit1t.excel.core.parse;

import com.github.stupdit1t.excel.callback.InCallback;
import com.github.stupdit1t.excel.common.ErrorMessage;
import com.github.stupdit1t.excel.common.PoiResult;
import com.github.stupdit1t.excel.common.PoiSheetDataArea;
import com.github.stupdit1t.excel.core.AbsParent;
import com.github.stupdit1t.excel.core.OpsPoiUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.function.Consumer;

/**
 * 导出规则定义
 */
public class OpsSheet<R> extends AbsParent<OpsParse<R>> {

    /**
     * 读取的sheetNum
     */
    int sheetIndex;

    /**
     * 读取的sheetName
     */
    String sheetName;

    /**
     * 选择的sheet模式
     * 1 下标
     * 2 名字
     */
    int sheetMode;

    /**
     * 表头行数量
     */
    int headerCount;

    /**
     * 尾部行数量
     */
    int footerCount;

    /**
     * 导入列定义
     */
    OpsColumn<R> opsColumn;

    /**
     * 行回调方法
     */
    InCallback<R> map;

    /**
     * 声明构造
     *
     * @param parent 当前对象
     */
    public OpsSheet(OpsParse<R> parent, int sheetIndex, int headerCount, int footerCount) {
        super(parent);
        this.headerCount = headerCount;
        this.footerCount = footerCount;
        checkSetSheetMode(1);
        this.sheetIndex = sheetIndex;
    }

    /**
     * 声明构造
     *
     * @param parent 当前对象
     */
    public OpsSheet(OpsParse<R> parent, String sheetName, int headerCount, int footerCount) {
        super(parent);
        this.headerCount = headerCount;
        this.footerCount = footerCount;
        checkSetSheetMode(2);
        this.sheetName = sheetName;
    }

    public OpsColumn<R> opsColumn() {
        return this.opsColumn(false);
    }

    /**
     * 列设置
     *
     * @param autoField 是否自动字段
     * @return
     */
    public OpsColumn<R> opsColumn(boolean autoField) {
        this.opsColumn = new OpsColumn<>(this, autoField);
        return this.opsColumn;
    }

    /**
     * 检测是否已经被设置状态
     */
    private void checkSetSheetMode(int wantSetMode) {
        if (sheetMode != 0 && sheetMode != wantSetMode) {
            throw new UnsupportedOperationException("仅支持设置 1 种sheet读取方式");
        }
        this.sheetMode = wantSetMode;
    }

    /**
     * 行回调方法
     *
     * @param map row 当前数据
     *            index 当前数据下标
     * @return OpsSheet
     */
    public OpsSheet<R> map(InCallback<R> map) {
        this.map = map;
        return this;
    }

    /**
     * 解析sheet方法
     *
     * @param partSize   批量页大小
     * @param partResult 批量结果
     */
    public void parsePart(int partSize, Consumer<PoiResult<R>> partResult) {
        try {
            // 校验用户输入, 非Map, 列必填
            if (this.opsColumn == null || (!parent.mapData && this.opsColumn.columns.isEmpty())) {
                throw new UnsupportedOperationException("导入的opsColumn字段不能为空!");
            }

            Map<String, InColumn<?>> columns =  this.opsColumn.columns;

            // 校验用户输入, 必填项校验
            if (StringUtils.isBlank(this.parent.fromPath) && this.parent.fromStream == null) {
                throw new UnsupportedOperationException("Excel来源不能为空!");
            }

            //1.根据excel报表获取OPCPackage
            OPCPackage opcPackage = null;
            if (this.parent.fromMode == 1) {
                opcPackage = OPCPackage.open(this.parent.fromPath, PackageAccess.READ);
            } else {
                opcPackage = OPCPackage.open(this.parent.fromStream);
            }
            //2.创建XSSFReader
            XSSFReader reader = new XSSFReader(opcPackage);
            //3.获取SharedStringTable对象
            SharedStrings table = reader.getSharedStringsTable();
            //4.获取styleTable对象
            StylesTable stylesTable = reader.getStylesTable();
            //5.创建Sax的xmlReader对象
            XMLReader xmlReader = XMLReaderFactory.createXMLReader();
            //6.注册事件处理器
            SheetHandler<R> sheetHandler = new SheetHandler<R>(this.opsColumn.autoField, this.sheetIndex, this.parent.rowClass, this.headerCount, columns, this.map, partSize, partResult, this.parent.allFields);
            XSSFSheetXMLHandler xmlHandler = new XSSFSheetXMLHandler(stylesTable, table, sheetHandler, false);
            xmlReader.setContentHandler(xmlHandler);
            //7.逐行读取
            XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) reader.getSheetsData();
            int index = 0;
            while (sheetIterator.hasNext()) {
                try (InputStream stream = sheetIterator.next()) {
                    if (index != this.sheetIndex) {
                        index++;
                        continue;
                    }
                    InputSource is = new InputSource(stream);
                    xmlReader.parse(is);
                    index++;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            partResult.accept(PoiResult.fail(new ErrorMessage(e)));
        }

    }

    /**
     * 解析sheet方法
     *
     * @return PoiResult
     */
    public PoiResult<R> parse() {
        // 校验用户输入, 非Map, 列必填
        if (this.opsColumn == null || (!parent.mapData && this.opsColumn.columns.isEmpty())) {
            throw new UnsupportedOperationException("导入的opsColumn字段不能为空!");
        }

        Map<String, InColumn<?>> columns = this.opsColumn.columns;
        // 校验用户输入, 必填项校验
        if (StringUtils.isBlank(this.parent.fromPath) && this.parent.fromStream == null) {
            throw new UnsupportedOperationException("Excel来源不能为空!");
        }


        PoiSheetDataArea poiSheetArea;
        if (StringUtils.isNotBlank(this.sheetName)) {
            poiSheetArea = new PoiSheetDataArea(this.sheetName, this.headerCount, this.footerCount);
        } else {
            poiSheetArea = new PoiSheetDataArea(this.sheetIndex, this.headerCount, this.footerCount);
        }
        if (this.parent.fromMode == 1) {
            if (this.parent.password != null) {
                return OpsPoiUtil.readSheet(this.parent.fromPath, this.parent.password, poiSheetArea, columns, this.map, this.parent.rowClass, this.parent.allFields, this.opsColumn.autoField);
            }
            return OpsPoiUtil.readSheet(this.parent.fromPath, poiSheetArea, columns, this.map, this.parent.rowClass, this.parent.allFields, this.opsColumn.autoField);
        } else if (this.parent.fromMode == 2) {
            if (this.parent.password != null) {
                return OpsPoiUtil.readSheet(this.parent.fromStream, this.parent.password, poiSheetArea, columns, this.map, this.parent.rowClass, this.parent.allFields, this.opsColumn.autoField);
            }
            return OpsPoiUtil.readSheet(this.parent.fromStream, poiSheetArea, columns, this.map, this.parent.rowClass, this.parent.allFields, this.opsColumn.autoField);
        }
        return PoiResult.fail(null);
    }
}