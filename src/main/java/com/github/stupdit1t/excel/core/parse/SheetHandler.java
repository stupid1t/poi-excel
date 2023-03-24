package com.github.stupdit1t.excel.core.parse;

import com.github.stupdit1t.excel.callback.InCallback;
import com.github.stupdit1t.excel.common.ErrorMessage;
import com.github.stupdit1t.excel.common.PoiCommon;
import com.github.stupdit1t.excel.common.PoiConstant;
import com.github.stupdit1t.excel.common.PoiResult;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

public class SheetHandler<T> implements XSSFSheetXMLHandler.SheetContentsHandler {

    private static final Logger LOG = LogManager.getLogger(SheetHandler.class);

    private final Class<T> entityClass;

    private final int sheetIndex;

    private final boolean mapClass;

    private T rowEntity;

    private final int headerRowNum;

    private final int batchSize;

    private final InCallback<T> callback;

    private final Map<String, InColumn<?>> columns;

    private int nowRowNum;

    private final Consumer<PoiResult<T>> partResult;

    private final List<T> data = new ArrayList<>();

    private final List<ErrorMessage> totalError = new ArrayList<>();

    private final List<ErrorMessage> rowError = new ArrayList<>();

    public SheetHandler(int sheetIndex, Class<T> entityClass, int headerRowNum, Map<String, InColumn<?>> columns, InCallback<T> callback, int batchSize, Consumer<PoiResult<T>> partResult) {
        this.sheetIndex = sheetIndex;
        this.entityClass = entityClass;
        this.mapClass = PoiCommon.isMapData(this.entityClass);
        this.headerRowNum = headerRowNum;
        this.columns = columns;
        this.callback = callback;
        this.batchSize = batchSize;
        this.partResult = partResult;
    }

    /**
     * 解析行开始
     */
    @Override
    public void startRow(int rowNum) {
        nowRowNum = rowNum;
        if (rowNum > headerRowNum - 1) {
            try {
                rowEntity = entityClass.newInstance();
            } catch (InstantiationException | IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 解析每一个单元格
     */
    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        if (this.nowRowNum < headerRowNum - 1) {
            return;
        }
        if (rowEntity != null) {
            CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf(cellReference);
            String columnIndexChar = PoiConstant.numsRefCell.get(cellRangeAddress.getFirstColumn());
            String location = columnIndexChar + cellRangeAddress.getFirstRow();
            try {
                InColumn<?> inColumn = this.columns.get(columnIndexChar);
                Object cellValue = formattedValue;
                String fieldName;
                if (inColumn != null) {
                    fieldName = inColumn.getField();
                } else {
                    // 只有map的情况下, 才使用列字符串
                    if (mapClass) {
                        fieldName = columnIndexChar;
                    } else {
                        fieldName = null;
                    }
                }
                if (fieldName == null) {
                    return;
                }

                // 校验类型转换处理
                if (inColumn != null) {
                    cellValue = inColumn.getCellVerifyRule().handle(cellValue);
                }

                if (mapClass) {
                    ((Map) rowEntity).put(fieldName, cellValue);
                } else {
                    FieldUtils.writeField(rowEntity, fieldName, cellValue, true);
                }
            } catch (Exception e) {
                rowError.add(new ErrorMessage(location, cellRangeAddress.getFirstRow(), cellRangeAddress.getFirstColumn(), e));
            }
        }
    }

    /**
     * 解析行结束
     */
    @Override
    public void endRow(int rowNum) {
        try {
            if (callback != null && (rowNum > headerRowNum - 1) && this.rowEntity != null) {
                callback.callback(this.rowEntity, rowNum);
            }
        } catch (Exception e) {
            LOG.error(e);
            rowError.add(new ErrorMessage("第" + (rowNum + 1) + "行", rowNum, -1, e));
        }

        if (rowError.isEmpty() && this.rowEntity != null) {
            data.add(this.rowEntity);
        } else {
            totalError.addAll(this.rowError);
        }

        if (data.size() == this.batchSize) {
            // 表示部分数据解析完, 结束
            batchFinish();
        }


    }

    /**
     * 部分批量结束
     */
    private void batchFinish() {
        PoiResult<T> poiResult = new PoiResult<>();
        poiResult.setData(data);
        poiResult.setError(rowError);
        if (!rowError.isEmpty()) {
            poiResult.setHasError(true);
        }
        partResult.accept(poiResult);
        data.clear();
        rowError.clear();
        totalError.clear();
    }

    @Override
    public void endSheet() {
        batchFinish();
    }

    //处理头尾
    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
    }
}