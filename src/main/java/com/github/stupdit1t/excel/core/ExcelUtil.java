package com.github.stupdit1t.excel.core;

import com.github.stupdit1t.excel.callback.InCallback;
import com.github.stupdit1t.excel.common.*;
import com.github.stupdit1t.excel.handle.ImgHandler;
import com.github.stupdit1t.excel.handle.rule.AbsCellVerifyRule;
import com.github.stupdit1t.excel.handle.rule.AbsSheetVerifyRule;
import com.github.stupdit1t.excel.handle.rule.CellVerifyRule;
import com.github.stupdit1t.excel.style.CellPosition;
import com.github.stupdit1t.excel.style.ICellStyle;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.util.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.text.ParseException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * excel导入导出工具，也可以导出模板
 *
 * @author 625
 */
public class ExcelUtil {

    private static final Logger LOG = LogManager.getLogger(ExcelUtil.class);

    /**
     * 私有
     */
    private ExcelUtil() {
    }

    /**
     * 导出入口
     *
     * @return OpsExport
     */
    public static OpsExport opsExport(WorkbookType workbookType) {
        return new OpsExport(workbookType);
    }

    /**
     * 设置打印方向
     *
     * @param sheet sheet页
     */
    static void printSetup(Sheet sheet) {
        PrintSetup printSetup = sheet.getPrintSetup();
        // 打印方向，true：横向，false：纵向
        printSetup.setLandscape(true);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);
    }

    /**
     * 给工作簿加密码 目前仅支持xlx
     *
     * @param workbook 工作簿
     * @param password 密码
     */
    static void encryptWorkbook(Workbook workbook, String password) {
        if (workbook instanceof HSSFWorkbook) {
            // 2003
            Biff8EncryptionKey.setCurrentUserPassword(password);
            ((HSSFWorkbook) workbook).writeProtectWorkbook(password, StringUtils.EMPTY);
        }
    }

    /**
     * 创建大数据workBook
     * 避免OOM,导出速度比较慢
     * <p>
     * 默认后缀 xlsx
     *
     * @param rowAccessWindowSize 在内存中的行数
     */
    static Workbook createBigWorkbook(int rowAccessWindowSize) {
        return new SXSSFWorkbook(rowAccessWindowSize);
    }

    /**
     * 创建空的workBook，做循环填充用
     *
     * @param xlsx 是否为xlsx格式
     */
    static Workbook createEmptyWorkbook(boolean xlsx) {
        Workbook wb;
        if (xlsx) {
            // 2007
            wb = new XSSFWorkbook();
        } else {
            // 2003
            wb = new HSSFWorkbook();
        }
        return wb;
    }

    /**
     * 创建空的workBook，做循环填充用
     *
     * @param xlsx 是否为xlsx格式
     */
    static Workbook createEmptyWorkbook(boolean xlsx, String password) {
        Workbook emptyWorkbook = createEmptyWorkbook(xlsx);
        encryptWorkbook(emptyWorkbook, password);
        return emptyWorkbook;
    }

    /**
     * 获取导出Excel的流
     *
     * @param response 响应流
     * @param fileName 文件名
     */
    static OutputStream getDownloadStream(HttpServletResponse response, String fileName) {
        try {
            if (fileName.endsWith(".xlsx")) {
                response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            } else {
                response.setContentType("application/vnd.ms-excel");
            }
            response.setCharacterEncoding(StandardCharsets.UTF_8.name());
            response.setHeader("Content-disposition", "attachment; filename=" + URLEncoder.encode(fileName, StandardCharsets.UTF_8.name()));
            return response.getOutputStream();
        } catch (IOException e) {
            LOG.error(e);
        }
        return null;
    }

    /**
     * 导出
     *
     * @param workbook 工作簿
     * @param response 响应
     * @param fileName 文件名
     */
    static void export(Workbook workbook, HttpServletResponse response, String fileName) {
        export(workbook, getDownloadStream(response, fileName));
    }

    /**
     * 导出
     *
     * @param out         导出流
     * @param data        数据源
     * @param exportRules 导出规则
     */
    static <T> void export(OutputStream out, List<T> data, ExportRules exportRules) {
        Workbook workbook = createEmptyWorkbook(exportRules.xlsx);
        if (StringUtils.isNotBlank(exportRules.password)) {
            encryptWorkbook(workbook, exportRules.password);
        }
        fillBook(workbook, data, exportRules);
        export(workbook, out);
    }

    /**
     * 导出
     *
     * @param workbook     工作簿
     * @param outputStream 流
     */
    static void export(Workbook workbook, OutputStream outputStream) {
        try (
                Workbook wb = workbook;
                OutputStream out = outputStream
        ) {
            wb.write(out);
        } catch (IOException e) {
            LOG.error(e);
        }
    }

    /**
     * 导出
     *
     * @param workbook 工作簿
     * @param outPath  删除目录
     */
    static void export(Workbook workbook, String outPath) {
        try (
                Workbook wb = workbook;
                OutputStream out = new FileOutputStream(outPath)
        ) {
            wb.write(out);
        } catch (IOException e) {
            LOG.error(e);
        }
    }

    /**
     * 填充wb，循环填充为多个Sheet
     *
     * @param wb          工作簿
     * @param data        数据
     * @param exportRules 导出规则
     */
    static <T> void fillBook(Workbook wb, List<T> data, ExportRules exportRules) {

        // -------------------- 全局样式处理 start ------------------------
        ICellStyle[] globalStyle = exportRules.globalStyle;
        // 标题样式设置
        Font titleFont = wb.createFont();
        CellStyle titleStyleSource = wb.createCellStyle();
        ICellStyle titleStyle = handleGlobalStyle(globalStyle, titleFont, titleStyleSource, CellPosition.TITLE);

        // 小标题样式
        Font headerFont = wb.createFont();
        CellStyle headerStyleSource = wb.createCellStyle();
        ICellStyle headerStyle = handleGlobalStyle(globalStyle, headerFont, headerStyleSource, CellPosition.HEADER);

        // 单元格样式
        Font cellFont = wb.createFont();
        CellStyle cellStyleSource = wb.createCellStyle();
        ICellStyle cellStyle = handleGlobalStyle(globalStyle, cellFont, cellStyleSource, CellPosition.CELL);

        // 尾行样式
        Font footerFont = wb.createFont();
        CellStyle footerStyleSource = wb.createCellStyle();
        ICellStyle footerStyle = handleGlobalStyle(globalStyle, footerFont, footerStyleSource, CellPosition.FOOTER);
        // -------------------- 全局样式处理 end  ------------------------

        String sheetName = exportRules.sheetName;
        Sheet sheet = sheetName != null ? wb.createSheet(sheetName) : wb.createSheet();
        ExcelUtil.printSetup(sheet);

        // ----------------------- 表头设置 start ------------------------

        // 创建表头
        for (int i = 0; i < exportRules.maxRows; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < exportRules.maxColumns; j++) {
                row.createCell(j);
            }
        }
        // 合并模式
        if (exportRules.ifMerge) {
            handleComplexHeader(exportRules, titleFont, titleStyleSource, titleStyle, headerFont, headerStyleSource, headerStyle, sheet);
        } else {// 非合并
            handleSimpleHeader(exportRules, titleFont, titleStyleSource, titleStyle, headerFont, headerStyleSource, headerStyle, sheet);
        }

        // ----------------------- 表头设置 end    ------------------------

        // ----------------------- 列属性设置 start -----------------------
        handleColumnProperty(data, exportRules, sheet);
        // ----------------------- 列属性设置 end  ------------------------

        // ----------------------- body设置 start ------------------------
        // 画图器
        @SuppressWarnings("unchecked") Drawing<Picture> createDrawingPatriarch = (Drawing<Picture>) sheet.createDrawingPatriarch();
        // 存储类的字段信息
        Map<Class<?>, Map<String, Field>> clsInfo = new HashMap<>();
        // 存储单元格样式信息，防止重复生成
        Map<String, CellStyle> cacheStyle = new HashMap<>();
        // 存储单元格字体信息，防止重复生成
        Map<String, Font> cacheFont = new HashMap<>();
        // 列信息
        List<Column<?>> fields = exportRules.column;
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + exportRules.maxRows);
            if (cellStyle.getHeight() != -1) {
                row.setHeight(cellStyle.getHeight());
            }
            // 行高自定义设置
            if (exportRules.cellHeight != -1) {
                row.setHeight(exportRules.cellHeight);
            }
            T t = data.get(i);
            for (int j = 0, n = 0; n < fields.size(); j++, n++) {
                Column<T> column = (Column<T>) fields.get(n);
                Cell cell = row.createCell(j);
                // 1.序号设置
                if (exportRules.autoNum && j == 0) {
                    cell.setCellValue(i + 1);
                    n--;
                    continue;
                }
                // 2.读取Map/Object对应字段值
                if (clsInfo.get(t.getClass()) == null) {
                    clsInfo.put(t.getClass(), PoiCommon.getAllFields(t.getClass()));
                }
                Object value = readField(clsInfo, t, column.field);

                // 3.填充列值
                Column.Style style = column.style;
                if (column.outHandle != null) {
                    style = Column.Style.clone(column.style);
                    value = column.outHandle.callback(value, t, style);
                }

                // 4.样式处理
                setCellStyle(wb, cellFont, cacheStyle, cacheFont, style, cell, value);

                // 5.设置单元格值
                setCellValue(createDrawingPatriarch, value, cell);

                // 6.批注添加
                String comment = column.comment;
                if (StringUtils.isNotBlank(comment)) {
                    // 表示需要用户添加批注
                    Drawing<?> drawingPatriarch = cell.getSheet().createDrawingPatriarch();
                    ClientAnchor clientAnchor;
                    RichTextString richTextString;
                    if (wb instanceof XSSFWorkbook) {
                        clientAnchor = new XSSFClientAnchor();
                        richTextString = new XSSFRichTextString(comment);
                    } else if (wb instanceof HSSFWorkbook) {
                        clientAnchor = new HSSFClientAnchor();
                        richTextString = new HSSFRichTextString(comment);
                    } else {
                        clientAnchor = new XSSFClientAnchor();
                        richTextString = new XSSFRichTextString(comment);
                    }
                    Comment cellComment = drawingPatriarch.createCellComment(clientAnchor);
                    cellComment.setAddress(cell.getAddress());
                    cellComment.setString(richTextString);
                    cell.setCellComment(cellComment);
                }
            }
        }
        // ----------------------- body设置 end     ------------------------

        // ------------------------footer设置 start  ------------------------
        handleFooter(data, exportRules, footerFont, footerStyleSource, footerStyle, sheet);
        // ------------------------footer设置 end ------------------------
    }

    /**
     * 处理单元格样式
     *
     * @param wb          工作簿
     * @param cellFont    原字体
     * @param cacheStyle  缓存样式
     * @param cacheFont   缓存字体
     * @param styleCustom 样式
     * @param cell        单元格
     * @param value       值
     * @param <T>
     */
    private static <T> void setCellStyle(Workbook wb, Font cellFont, Map<String, CellStyle> cacheStyle, Map<String, Font> cacheFont, Column.Style styleCustom, Cell cell, Object value) {
        String styleCacheKey = styleCustom.getStyleCacheKey();
        // 此处有值, 表示用户自定义列样式
        if (styleCacheKey != null) {
            CellStyle style = cacheStyle.get(styleCacheKey);
            // 表示缓存无, 重新构建
            if (style == null) {
                style = wb.createCellStyle();
                style.cloneStyleFrom(cell.getCellStyle());

                // 1.水平定位
                HorizontalAlignment align = styleCustom.align;
                if (align != null) {
                    style.setAlignment(align);
                }

                // 2.垂直定位
                VerticalAlignment valign = styleCustom.valign;
                if (valign != null) {
                    style.setVerticalAlignment(valign);
                }

                // 3.字体颜色
                IndexedColors color = styleCustom.color;
                if (color != null) {
                    Font font = cacheFont.get(styleCacheKey);
                    if (font == null) {
                        font = wb.createFont();
                        PoiCommon.copyFont(font, cellFont);
                        cacheFont.put(styleCacheKey, font);
                    }
                    font.setColor(color.getIndex());
                    style.setFont(font);
                }

                // 4.背景色
                IndexedColors backColor = styleCustom.backColor;
                if (backColor != null) {
                    style.setFillForegroundColor(backColor.getIndex());
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }

                // 5. 日期格式化
                String pattern = styleCustom.datePattern;
                if (StringUtils.isNotBlank(pattern) && (value instanceof Date || value instanceof LocalDate || value instanceof LocalDateTime)) {
                    CreationHelper createHelper = wb.getCreationHelper();
                    style.setDataFormat(createHelper.createDataFormat().getFormat(pattern));
                }
                cacheStyle.put(styleCacheKey, style);
            }
            // 最终样式设置
            cell.setCellStyle(style);
        }

        // 6.高度
        int height = styleCustom.height;
        if (height != -1) {
            // 表示需要用户自定义高度
            cell.getRow().setHeight((short) height);
        }
    }

    /**
     * footer设置
     *
     * @param data              数据
     * @param exportRules       导出规则
     * @param footerFont        font
     * @param footerStyleSource style
     * @param sheet             sheet
     * @param footerStyle       全局样式
     */
    private static <T> void handleFooter(List<T> data, ExportRules exportRules, Font footerFont, CellStyle footerStyleSource, ICellStyle footerStyle, Sheet sheet) {
        if (exportRules.ifFooter) {
            Workbook workbook = sheet.getWorkbook();
            List<ComplexCell> footerRules = exportRules.footerRules;
            // 构建尾行
            int currRowNum = exportRules.maxRows + data.size();
            int[] footerNum = getFooterNum(footerRules, currRowNum);
            for (int j : footerNum) {
                sheet.createRow(j);
            }

            for (ComplexCell entry : footerRules) {
                Integer[] range = entry.getLocationIndex();
                String value = entry.getText();
                BiConsumer<Font, CellStyle> fontCellStyle = entry.getStyle();
                int firstRow = range[0] + currRowNum;
                int lastRow = range[1] + currRowNum;
                int firstCol = range[2];
                int lastCol = range[3];
                CellStyle styleNew;
                if (fontCellStyle != null) {
                    // 自定义header单元格样式
                    styleNew = workbook.createCellStyle();
                    Font fontNew = workbook.createFont();
                    PoiCommon.copyStyleAndFont(styleNew, fontNew, footerStyleSource, footerFont);
                    fontCellStyle.accept(fontNew, styleNew);
                } else {
                    styleNew = footerStyleSource;
                }
                Cell cell = CellUtil.createCell(sheet.getRow(firstRow), firstCol, value, styleNew);
                if (value.startsWith("=")) {
                    cell.setCellFormula(value.substring(1));
                }
                if ((lastRow - firstRow) != 0 || (lastCol - firstCol) != 0) {
                    cellMerge(sheet, firstRow, lastRow, firstCol, lastCol);
                }
                // 行高自定义设置
                for (int i = firstRow; i <= lastRow; i++) {
                    if (footerStyle.getHeight() != -1) {
                        sheet.getRow(i).setHeight(footerStyle.getHeight());
                    }
                    if (exportRules.footerHeight != -1) {
                        sheet.getRow(i).setHeight(exportRules.footerHeight);
                    }
                }
            }
        }
    }

    /**
     * 列属性设置
     *
     * @param data        数据
     * @param exportRules 导出规则
     * @param sheet       sheet
     */
    private static <T> void handleColumnProperty(List<T> data, ExportRules exportRules, Sheet sheet) {
        // ----------------------- 列属性设置 start--------------------
        List<Column<?>> fields = exportRules.column;
        int autoNumColumnWidth = exportRules.autoNumColumnWidth;
        for (int i = 0, j = 0; i < fields.size(); i++, j++) {
            // 0.每一列默认单元格样式设置
            // 1.width设置
            if (exportRules.autoNum && j == 0) {
                j++;
                sheet.setColumnWidth(0, autoNumColumnWidth);
            }
            Column<?> column = fields.get(i);
            // 1.1是否自动列宽
            int width = column.style.width;
            if (width != -1) {
                sheet.setColumnWidth(j, width);
            } else {
                try {
                    // 1.2根据maxRows，获取表头的值设置宽度
                    Row row = sheet.getRow(exportRules.maxRows - 1);
                    String headerValue = row.getCell(j).getStringCellValue();
                    if (StringUtils.isBlank(headerValue)) {
                        row = sheet.getRow(exportRules.maxRows - 2);
                        headerValue = row.getCell(j).getStringCellValue();
                    }
                    sheet.setColumnWidth(j, headerValue.getBytes().length * 256);
                } catch (Exception e) {
                    if (exportRules.autoNum) {
                        throw new UnsupportedOperationException("自动序号设置错误，请确认在header添加序号列");
                    } else {
                        throw e;
                    }
                }
            }
            // 2.downDown设置
            int lastRow = (exportRules.maxRows - 1) + data.size();
            lastRow = lastRow == (exportRules.maxRows - 1) ? PoiConstant.MAX_FILL_COL : lastRow;
            String[] dropdown = column.dropdown;
            if (dropdown != null && dropdown.length > 0) {
                sheet.addValidationData(createDropDownValidation(sheet, dropdown, j, exportRules.maxRows, lastRow));
            }

            // 3.时间校验
            String date = column.verifyDate;
            if (date != null) {
                String[] split = date.split("@");
                String info = null;
                if (split.length == 2) {
                    info = split[1];
                }
                String[] split1 = split[0].split("~");
                if (split1.length < 2) {
                    throw new IllegalArgumentException("时间校验表达式不正确,请填写如" + column.style.datePattern + "的值!");
                }
                try {
                    sheet.addValidationData(createDateValidation(sheet, column.style.datePattern, split1[0], split1[1], info, j, exportRules.maxRows, lastRow));
                } catch (ParseException e) {
                    LOG.error(e);
                    throw new IllegalArgumentException("时间校验表达式不正确,请填写如" + column.style.datePattern + "的值!");
                } catch (Exception e) {
                    LOG.error(e);
                }
            }

            // 4.整数数字校验
            String num = column.verifyIntNum;
            if (num != null) {
                String[] split = num.split("@");
                String info = null;
                if (split.length == 2) {
                    info = split[1];
                }
                String[] split1 = split[0].split("~");
                if (split1.length < 2) {
                    throw new IllegalArgumentException("数字表达式不正确,请填写如10~30的值!");
                }
                sheet.addValidationData(createNumValidation(sheet, split1[0], split1[1], info, j, exportRules.maxRows, lastRow));
            }

            // 4.浮点数字校验
            String floatNum = column.verifyFloatNum;
            if (floatNum != null) {
                String[] split = floatNum.split("@");
                String info = null;
                if (split.length == 2) {
                    info = split[1];
                }
                String[] split1 = split[0].split("~");
                if (split1.length < 2) {
                    throw new IllegalArgumentException("数字表达式不正确,请填写如10.0~30.0的值!");
                }
                sheet.addValidationData(createFloatValidation(sheet, split1[0], split1[1], info, j, exportRules.maxRows, lastRow));
            }

            // 5.自定义校验
            String custom = column.verifyCustom;
            if (custom != null) {
                String[] split = custom.split("@");
                String info = null;
                if (split.length == 2) {
                    info = split[1];
                }
                sheet.addValidationData(createCustomValidation(sheet, split[0], info, j, exportRules.maxRows, lastRow));
            }

            // 6.文本长度校验
            String text = column.verifyText;
            if (text != null) {
                String[] split1 = text.split("@");
                String info = null;
                if (split1.length == 2) {
                    info = split1[1];
                }
                String[] split2 = split1[0].split("~");
                if (split2.length < 2) {
                    throw new IllegalArgumentException("文本长度校验格式不正确，请设置如3~10格式!");
                }
                sheet.addValidationData(createTextLengthValidation(sheet, split2[0], split2[1], info, j, exportRules.maxRows, lastRow));
            }
        }
    }

    /**
     * 简单表头设计
     *
     * @param exportRules       导出规则
     * @param titleFont         大标题字体
     * @param titleStyleSource  大标题样式
     * @param titleStyle        大标题自定义样式处理
     * @param headerFont        标题字体
     * @param headerStyleSource 标题样式
     * @param headerStyle       标题自定义样式处理
     * @param sheet             sheet
     */
    private static void handleSimpleHeader(ExportRules exportRules, Font titleFont, CellStyle titleStyleSource, ICellStyle titleStyle, Font headerFont, CellStyle headerStyleSource, ICellStyle headerStyle, Sheet sheet) {
        // 1. 冻结表头
        if (exportRules.freezeHeader) {
            sheet.createFreezePane(0, exportRules.maxRows, 0, exportRules.maxRows);
        }

        // 2. title 内容设置和行高
        if (exportRules.title != null) {
            // title全局行高
            if (titleStyle.getHeight() != -1) {
                sheet.getRow(0).setHeight(titleStyle.getHeight());
            }
            // title行高自定义设置
            if (exportRules.titleHeight != -1) {
                sheet.getRow(0).setHeight(exportRules.titleHeight);
            }
            CellUtil.createCell(sheet.getRow(0), 0, exportRules.title, titleStyleSource);
        }

        // 3.header设置和行高
        int headerIndex = exportRules.title == null ? 0 : 1;
        // header全局行高
        if (headerStyle.getHeight() != -1) {
            sheet.getRow(headerIndex).setHeight(headerStyle.getHeight());
        }
        // header行高自定义设置
        if (exportRules.headerHeight != -1) {
            sheet.getRow(headerIndex).setHeight(exportRules.headerHeight);
        }
        LinkedHashMap<String, BiConsumer<Font, CellStyle>> headerMap = exportRules.simpleHeader;
        List<String> header = new ArrayList<>(headerMap.keySet());
        for (int i = 0; i < header.size(); i++) {
            String text = header.get(i);
            CellStyle styleNew;
            BiConsumer<Font, CellStyle> fontCellStyle = headerMap.get(text);
            if (fontCellStyle != null) {
                // 自定义header单元格样式
                styleNew = sheet.getWorkbook().createCellStyle();
                Font fontNew = sheet.getWorkbook().createFont();
                PoiCommon.copyStyleAndFont(styleNew, fontNew, headerStyleSource, headerFont);
                fontCellStyle.accept(fontNew, styleNew);
            } else {
                styleNew = headerStyleSource;
            }
            CellUtil.createCell(sheet.getRow(headerIndex), i, text, styleNew);
        }
    }

    /**
     * 复杂表头设计
     *
     * @param exportRules       导出规则
     * @param titleFont         大标题字体
     * @param titleStyleSource  大标题样式
     * @param titleStyle        大标题自定义样式处理
     * @param headerFont        标题字体
     * @param headerStyleSource 标题样式
     * @param headerStyle       标题自定义样式处理
     * @param sheet             sheet
     */
    private static void handleComplexHeader(ExportRules exportRules, Font titleFont, CellStyle titleStyleSource, ICellStyle titleStyle, Font headerFont, CellStyle headerStyleSource, ICellStyle headerStyle, Sheet sheet) {
        // 冻结表头
        if (exportRules.freezeHeader) {
            sheet.createFreezePane(0, exportRules.maxRows, 0, exportRules.maxRows);
        }
        // header
        List<ComplexCell> complexHeader = exportRules.complexHeader;
        for (ComplexCell complexCell : complexHeader) {
            Integer[] range = complexCell.getLocationIndex();
            // 合并表头
            int firstRow = range[0];
            int lastRow = range[1];
            int firstCol = range[2];
            int lastCol = range[3];
            CellStyle styleTemp;
            Font fontTemp;
            if ((exportRules.maxColumns - 1) == lastCol - firstCol && firstRow == 0) {
                // 占满全格, 且第一行开始为表头
                for (int i = firstRow; i <= lastRow; i++) {
                    if (titleStyle.getHeight() != -1) {
                        sheet.getRow(i).setHeight(titleStyle.getHeight());
                    }
                    // 行高自定义设置
                    if (exportRules.titleHeight != -1) {
                        sheet.getRow(i).setHeight(exportRules.titleHeight);
                    }
                }
                styleTemp = titleStyleSource;
                fontTemp = titleFont;
            } else {
                // 没有大表头, 普通合并格
                for (int i = firstRow; i <= lastRow; i++) {
                    if (headerStyle.getHeight() != -1) {
                        sheet.getRow(i).setHeight(headerStyle.getHeight());
                    }
                    // 行高自定义设置
                    if (exportRules.headerHeight != -1) {
                        sheet.getRow(i).setHeight(exportRules.headerHeight);
                    }
                }
                styleTemp = headerStyleSource;
                fontTemp = headerFont;
            }
            CellStyle styleNew = styleTemp;
            BiConsumer<Font, CellStyle> fontCellStyle = complexCell.getStyle();
            if (fontCellStyle != null) {
                // 自定义header单元格样式
                styleNew = sheet.getWorkbook().createCellStyle();
                Font fontNew = sheet.getWorkbook().createFont();
                PoiCommon.copyStyleAndFont(styleNew, fontNew, styleTemp, fontTemp);
                fontCellStyle.accept(fontNew, styleNew);
            }
            CellUtil.createCell(sheet.getRow(firstRow), firstCol, complexCell.getText(), styleNew);
            if ((lastRow - firstRow) != 0 || (lastCol - firstCol) != 0) {
                cellMerge(sheet, firstRow, lastRow, firstCol, lastCol);
            }
        }
    }

    /**
     * 合并单元格
     *
     * @param sheet    sheet
     * @param firstRow 卡死行
     * @param lastRow  结束行
     * @param firstCol 开始列
     * @param lastCol  结束列
     */
    private static void cellMerge(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        CellRangeAddress cra = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(cra);
        RegionUtil.setBorderTop(BorderStyle.THIN, cra, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, cra, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, cra, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, cra, sheet);
    }

    /**
     * 全局样式处理
     *
     * @param globalStyle  全局样式
     * @param font         字体
     * @param cellStyle    样式
     * @param cellPosition 位置
     */
    private static ICellStyle handleGlobalStyle(ICellStyle[] globalStyle, Font font, CellStyle cellStyle, CellPosition cellPosition) {
        ICellStyle titleStyle = ICellStyle.getCellStyleByPosition(cellPosition, globalStyle);
        cellStyle.setFont(font);
        titleStyle.handleStyle(font, cellStyle);
        return titleStyle;
    }

    /**
     * 解析Sheet
     *
     * @param cls                结果bean
     * @param absSheetVerifyRule 校验器
     * @param sheet              解析的sheet
     * @param dataStartRow       开始行:从0开始计，表示excel第一行
     * @param dataEndRowCount    尾行非数据行数量，比如统计行2行，则写2
     * @param callback           加入回调逻辑
     * @return ImportRspInfo
     */
    static <T> PoiResult<T> readSheet(Sheet sheet, Class<T> cls, Consumer<AbsSheetVerifyRule> absSheetVerifyRule, int dataStartRow, int dataEndRowCount, InCallback<T> callback) {
        // 公式计算初始化
        FormulaEvaluator formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
        AbsSheetVerifyRule verifyBuilder = AbsSheetVerifyRule.buildRule(absSheetVerifyRule);
        // 规则初始化
        verifyBuilder.init();
        PoiResult<T> rsp = new PoiResult<>();
        List<T> beans = new ArrayList<>();
        // 获取excel中所有图片
        List<String> imgField = new ArrayList<>();
        Map<String, PictureData> pictures = null;
        Map<String, CellVerifyRule> verifies = verifyBuilder.getColumnVerifyRule();
        Set<String> keySet = verifies.keySet();
        int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
        for (String key : keySet) {
            CellVerifyRule cellVerifyRule = verifies.get(key);
            AbsCellVerifyRule<?> cellVerify = cellVerifyRule.getCellVerify();
            if (cellVerify instanceof ImgHandler) {
                imgField.add(key);
                if (pictures == null || pictures.isEmpty()) {
                    pictures = getSheetPictures(sheetIndex, sheet);
                }
            }
        }
        StringBuilder errors = new StringBuilder();
        StringBuilder rowErrors = new StringBuilder();
        try {
            int rowStart = sheet.getFirstRowNum() + dataStartRow;
            // warn获取真实的数据行尾数
            int rowEnd = getLastRealLastRow(sheet.getRow(sheet.getLastRowNum())) - dataEndRowCount;
            for (int rowNum = rowStart; rowNum <= rowEnd; rowNum++) {
                Row r = sheet.getRow(rowNum);
                if (r == null) {
                    continue;
                }
                // 创建对象
                T t = cls.newInstance();
                int fieldNum = 0;
                String[] cellRefs = verifyBuilder.getCellRefs();
                for (String index : cellRefs) {
                    // 列坐标
                    Integer cellNum = PoiConstant.cellRefNums.get(index);
                    CellReference cellRef = new CellReference(rowNum, cellNum);
                    String filed = verifyBuilder.getFields()[fieldNum];
                    try {
                        Object cellValue;
                        if (imgField.size() > 0 && imgField.contains(filed)) {
                            String pictureIndex = sheetIndex + "," + rowNum + "," + cellNum;
                            PictureData remove = pictures.remove(pictureIndex);
                            cellValue = remove == null ? null : remove.getData();
                        } else {
                            cellValue = getCellValue(r, cellNum, formulaEvaluator);
                        }
                        // 校验和格式化列值
                        cellValue = verifyBuilder.verify(filed, cellValue);
                        // 填充列值
                        FieldUtils.writeField(t, filed, cellValue, true);
                    } catch (PoiException e) {
                        rowErrors.append("[").append(cellRef.formatAsString()).append("]").append(e.getMessage()).append("\t");
                    }
                    fieldNum++;
                }
                // 回调处理一下特殊逻辑
                if (callback != null) {
                    try {
                        callback.callback(t, rowNum);
                    } catch (PoiException e) {
                        rowErrors.append(e.getMessage()).append("\t");
                    }
                }
                beans.add(t);
                if (rowErrors.length() > 0) {
                    errors.append(rowErrors).append("\r\n");
                    rowErrors.setLength(0);
                }
            }
        } catch (Exception e) {
            LOG.error(e);
        } finally {
            // throw parse exception
            if (errors.length() > 0) {
                rsp.setSuccess(false);
                rsp.setMessage(errors.toString());
            }
            rsp.setData(beans);
        }
        // 返回结果
        return rsp;
    }

    /**
     * 读取规则excel数据内容为map
     *
     * @param filePath        文件路径
     * @param sheetNum        表格号
     * @param dataStartRow    开始读取行
     * @param dataEndRowCount 尾部
     * @return List<Map < String, Object>>
     */
    static List<Map<String, Object>> readSheet(String filePath, int sheetNum, int dataStartRow, int dataEndRowCount) {
        try (InputStream is = new FileInputStream(filePath)) {
            return readSheet(is, sheetNum, dataStartRow, dataEndRowCount);
        } catch (IOException e) {
            LOG.error(e);
        }
        return Collections.emptyList();
    }

    /**
     * 读取规则excel数据内容为map
     *
     * @param is              文件流
     * @param dataStartRow    数据起始行
     * @param dataEndRowCount 尾部非数据行数量
     * @return List<Map < String, Object>>
     */
    static List<Map<String, Object>> readSheet(InputStream is, int sheetNum, int dataStartRow, int dataEndRowCount) {
        try (Workbook wb = WorkbookFactory.create(is)) {
            Sheet sheet = wb.getSheetAt(sheetNum);
            return readSheet(sheet, dataStartRow, dataEndRowCount);
        } catch (Exception e) {
            LOG.error(e);
        }
        return Collections.emptyList();
    }

    /**
     * 读取规则excel数据内容为map
     *
     * @param filePath        文件
     * @param dataStartRow    数据起始行
     * @param dataEndRowCount 尾部非数据行数量
     * @return List<Map < String, Object>>
     */
    static <T> PoiResult<T> readSheet(String filePath, Class<T> cls, Consumer<AbsSheetVerifyRule> absSheetVerifyRule, int sheetNum, int dataStartRow, int dataEndRowCount, InCallback<T> callback) {
        try (InputStream is = new FileInputStream(filePath)) {
            return readSheet(is, cls, absSheetVerifyRule, sheetNum, dataStartRow, dataEndRowCount, callback);
        } catch (IOException e) {
            LOG.error(e);
        }
        return PoiResult.fail();
    }

    /**
     * 读取规则excel数据内容为map
     *
     * @param is              文件流
     * @param dataStartRow    数据起始行
     * @param dataEndRowCount 尾部非数据行数量
     * @return List<Map < String, Object>>
     */
    static <T> PoiResult<T> readSheet(InputStream is, Class<T> cls, Consumer<AbsSheetVerifyRule> absSheetVerifyRule, int sheetNum, int dataStartRow, int dataEndRowCount, InCallback<T> callback) {
        try (Workbook wb = WorkbookFactory.create(is)) {
            Sheet sheet = wb.getSheetAt(sheetNum);
            return readSheet(sheet, cls, absSheetVerifyRule, dataStartRow, dataEndRowCount, callback);
        } catch (Exception e) {
            LOG.error(e);
        }
        return PoiResult.fail();
    }

    /**
     * 读取规则excel数据内容为map
     *
     * @param filePath        文件
     * @param dataStartRow    数据起始行
     * @param dataEndRowCount 尾部非数据行数量
     * @return List<Map < String, Object>>
     */
    static <T> PoiResult<T> readSheet(String filePath, Class<T> cls, Consumer<AbsSheetVerifyRule> absSheetVerifyRule, int sheetNum, int dataStartRow, int dataEndRowCount) {
        try (InputStream is = new FileInputStream(filePath)) {
            return readSheet(is, cls, absSheetVerifyRule, sheetNum, dataStartRow, dataEndRowCount);
        } catch (IOException e) {
            LOG.error(e);
        }
        return PoiResult.fail();
    }

    /**
     * 读取规则excel数据内容为map
     *
     * @param is              文件流
     * @param dataStartRow    数据起始行
     * @param dataEndRowCount 尾部非数据行数量
     * @return List<Map < String, Object>>
     */
    static <T> PoiResult<T> readSheet(InputStream is, Class<T> cls, Consumer<AbsSheetVerifyRule> absSheetVerifyRule, int sheetNum, int dataStartRow, int dataEndRowCount) {
        try (Workbook wb = WorkbookFactory.create(is)) {
            Sheet sheet = wb.getSheetAt(sheetNum);
            return readSheet(sheet, cls, absSheetVerifyRule, dataStartRow, dataEndRowCount, null);
        } catch (Exception e) {
            LOG.error(e);
        }
        return PoiResult.fail();
    }

    /**
     * 读取规则excel数据内容为map
     *
     * @param sheet           sheet页
     * @param dataStartRow    起始行
     * @param dataEndRowCount 尾部非数据行数量
     * @return List<Map < String, Object>>
     */
    static List<Map<String, Object>> readSheet(Sheet sheet, int dataStartRow, int dataEndRowCount) {
        List<Map<String, Object>> sheetData = new ArrayList<>();
        int rowStart = sheet.getFirstRowNum() + dataStartRow;
        // 获取真实的数据行尾数
        int rowEnd = getLastRealLastRow(sheet.getRow(sheet.getLastRowNum())) - dataEndRowCount;
        FormulaEvaluator formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
        for (int j = rowStart; j <= rowEnd; j++) {
            Map<String, Object> cellMap = new HashMap<>();
            Row row = sheet.getRow(j);
            if (row == null) {
                continue;
            }
            short lastCellNum = row.getLastCellNum();
            for (int k = 0; k < lastCellNum; k++) {
                Object cellValue = getCellValue(row, k, formulaEvaluator);
                cellMap.put(PoiConstant.numsRefCell.get(k), cellValue);
            }
            sheetData.add(cellMap);
        }
        // 返回结果
        return sheetData;
    }

    /**
     * 读取excel,替换内置变量
     *
     * @param filePath 文件路径
     * @param variable 内置变量
     */
    static Workbook readExcelWrite(String filePath, Map<String, String> variable) {
        try (FileInputStream is = new FileInputStream(filePath)) {

            return readExcelWrite(is, variable);
        } catch (IOException e) {
            LOG.error(e);
        }
        return null;
    }

    /**
     * 读取excel,替换内置变量
     *
     * @param is       excel文件流
     * @param variable 内置变量
     */
    static Workbook readExcelWrite(InputStream is, Map<String, String> variable) {
        try {
            Workbook wb = WorkbookFactory.create(is);
            return readExcelWrite(wb, variable);
        } catch (IOException e) {
            LOG.error(e);
        }
        return null;
    }

    /**
     * 读取excel,替换内置变量
     *
     * @param workbook excel对象
     * @param variable 内置变量
     */
    private static Workbook readExcelWrite(Workbook workbook, Map<String, String> variable) {
        int numberOfSheets = workbook.getNumberOfSheets();
        FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            Row lastRow = sheet.getRow(sheet.getLastRowNum());
            int lastRealLastRow = getLastRealLastRow(lastRow);
            for (int j = 0; j <= lastRealLastRow; j++) {
                Row row = sheet.getRow(j);
                if (row == null) {
                    continue;
                }
                short lastCellNum = row.getLastCellNum();
                for (short k = 0; k < lastCellNum; k++) {
                    Object cellValue = getCellValue(row, k, formulaEvaluator);
                    if (cellValue instanceof String) {
                        String cellValueStr = (String) cellValue;
                        if (!cellValueStr.contains("$")) {
                            continue;
                        }
                        Set<String> vars = variable.keySet();
                        for (String var : vars) {
                            String value = variable.get(var);
                            cellValueStr = cellValueStr.replace("${" + var + "}", value);
                        }
                        if (cellValueStr.startsWith("=")) {
                            row.getCell(k).setCellFormula(cellValueStr.substring(1));
                        } else {
                            row.getCell(k).setCellValue(cellValueStr);
                        }
                    }
                }
            }
        }
        return workbook;
    }

    /**
     * 获取真实的数据行
     *
     * @param row 单元格
     * @return int
     */
    private static int getLastRealLastRow(Row row) {
        Sheet sheet = row.getSheet();
        short lastCellNum = row.getLastCellNum();
        if (lastCellNum == -1) {
            int rowNum = row.getRowNum();
            Row newRow = sheet.getRow(--rowNum);
            while (newRow == null) {
                newRow = sheet.getRow(--rowNum);
            }
            return getLastRealLastRow(newRow);
        } else {
            int blankCell = 0;
            for (int i = 0; i < lastCellNum; i++) {
                Cell cell = row.getCell(i);
                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    blankCell++;
                }
            }
            if (blankCell >= lastCellNum) {
                int rowNum = row.getRowNum();
                Row newRow = sheet.getRow(--rowNum);
                while (newRow == null) {
                    newRow = sheet.getRow(--rowNum);
                }
                return getLastRealLastRow(newRow);
            }
        }
        return row.getRowNum();
    }

    /**
     * 读取字段的值
     *
     * @param clsInfo 类信息
     * @param t       当前值
     * @param fields  字段名称
     * @return Object
     */
    private static Object readField(Map<Class<?>, Map<String, Field>> clsInfo, Object t, String fields) {
        // 读取子属性
        String[] split = fields.split("\\.");
        Object value = t;
        for (int i = 0; i < split.length; i++) {
            value = readObjectFieldValue(value, split[i], clsInfo);
            // 属性为空跳出
            if (value == null) {
                return "";
            }
            if (i == split.length - 1) {
                return value;
            }
        }
        return "";
    }

    /**
     * 读取object的属性
     *
     * @param t       对象
     * @param key     field字段
     * @param clsInfo 类信息
     * @return Object
     */
    private static Object readObjectFieldValue(Object t, String key, Map<Class<?>, Map<String, Field>> clsInfo) {
        try {
            if (t instanceof Map) {
                t = ((Map<?, ?>) t).get(key);
            } else {
                Class<?> subCls = t.getClass();
                Map<String, Field> subField = clsInfo.get(subCls);
                if (subField == null) {
                    subField = PoiCommon.getAllFields(subCls);
                    clsInfo.put(subCls, subField);
                }
                Field field = subField.get(key);
                if (field == null) {
                    // 为方法，不是属性
                    char[] charName = key.toCharArray();
                    charName[0] -= 32;
                    String methodName = "get" + String.valueOf(charName);
                    Method method = subCls.getMethod(methodName);
                    t = method.invoke(t);
                } else {
                    t = field.get(t);
                }
            }
        } catch (Exception e) {
            LOG.error(e);
            t = null;
        }
        return t;
    }


    /**
     * 给单元格设置值
     *
     * @param createDrawingPatriarch 画图器
     * @param value                  单元格值
     * @param cell                   单元格
     */
    private static void setCellValue(Drawing<Picture> createDrawingPatriarch, Object value, Cell cell) {
        Workbook workbook = cell.getSheet().getWorkbook();

        // 8.值设置, 判断值的类型后进行强制类型转换.再设置单元格格式
        if (value instanceof String) {
            // 判断是否是公式
            String strValue = String.valueOf(value);
            if (strValue.startsWith("=")) {
                cell.setCellFormula(strValue.substring(1));
            } else {
                cell.setCellValue(strValue);
            }
        } else if (value instanceof Number) {
            // 处理整形自动不展示小数点
            if (value instanceof Integer || value instanceof Long || value instanceof Short) {
                cell.setCellValue(String.valueOf(value));
            } else {
                cell.setCellValue(((Number) value).doubleValue());
            }
        } else if (value instanceof Date || value instanceof LocalDate || value instanceof LocalDateTime) {
            if (value instanceof Date) {
                Date date = (Date) value;
                cell.setCellValue(date);
            } else if (value instanceof LocalDateTime) {
                LocalDateTime date = (LocalDateTime) value;
                cell.setCellValue(date);
            } else {
                LocalDate date = (LocalDate) value;
                cell.setCellValue(date);
            }
        } else if (value instanceof byte[]) {
            byte[] data = (byte[]) value;
            // 5.1anchor主要用于设置图片的属性
            short x = (short) cell.getColumnIndex();
            int y = cell.getRowIndex();
            // 5.2插入图片
            ClientAnchor anchor;
            int add1;
            if (workbook instanceof XSSFWorkbook) {
                anchor = new XSSFClientAnchor(20, 20, 20, 20, x, y, x + 1, y + 1);
                add1 = workbook.addPicture(data, XSSFWorkbook.PICTURE_TYPE_PNG);
            } else if (workbook instanceof HSSFWorkbook) {
                anchor = new HSSFClientAnchor(20, 20, 20, 20, x, y, (short) (x + 1), y + 1);
                add1 = workbook.addPicture(data, SXSSFWorkbook.PICTURE_TYPE_PNG);
            } else {
                anchor = new XSSFClientAnchor(20, 20, 20, 20, x, y, (short) (x + 1), y + 1);
                add1 = workbook.addPicture(data, XSSFWorkbook.PICTURE_TYPE_PNG);
            }
            createDrawingPatriarch.createPicture(anchor, add1);
            cell.setCellValue("");
        } else if (value == null) {
            cell.setCellValue("");
        } else {
            cell.setCellValue(String.valueOf(value));
        }
    }


    /**
     * 根据页脚数据获得行号
     *
     * @param entries    规则
     * @param currRowNum 当前行
     * @return int[]
     */
    private static int[] getFooterNum(List<ComplexCell> entries, int currRowNum) {
        int row = 0;
        List<Integer[]> rules = entries.stream().map(ComplexCell::getLocationIndex).collect(Collectors.toList());
        for (Integer[] range : rules) {
            int a = range[1] + 1;
            row = Math.max(a, row);
        }
        int[] footerNum = new int[row];
        for (int i = 0; i < row; i++) {
            footerNum[i] = currRowNum + i;
        }
        return footerNum;
    }

    /**
     * 获取单元格的值
     *
     * @param r       当前行
     * @param cellNum 单元格号
     * @return Object
     */
    private static Object getCellValue(Row r, int cellNum, FormulaEvaluator formulaEvaluator) {
        // 缺失列处理政策
        Cell cell = r.getCell(cellNum, MissingCellPolicy.CREATE_NULL_AS_BLANK);
        Object obj = null;
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case STRING:
                obj = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    obj = cell.getDateCellValue();
                } else {
                    obj = cell.getNumericCellValue();
                }
                break;
            case BOOLEAN:
                obj = cell.getBooleanCellValue();
                break;
            case FORMULA:
                // 拿到计算公式eval
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                switch (evaluate.getCellType()) {
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            obj = cell.getDateCellValue();
                        } else {
                            obj = cell.getNumericCellValue();
                        }
                        break;
                    case STRING:
                        obj = evaluate.getStringValue();
                        break;
                    default:
                        obj = cell.getRichStringCellValue().getString();
                }
                break;
            case BLANK:
                obj = "";
                break;
            default:
                break;
        }
        return obj;
    }

    /**
     * 获取Excel2003图片
     *
     * @param sheetNum 当前sheet下标
     * @param sheet    当前sheet对象
     * @return Map key:图片单元格索引（0-sheet下标,1-列号,1-行号）String，value:图片流PictureData
     */
    private static Map<String, PictureData> getSheetPictures(int sheetNum, Sheet sheet) {
        if (sheet instanceof HSSFSheet) {
            HSSFSheet sheetHssf = (HSSFSheet) sheet;
            return getSheetPictures03(sheetNum, sheetHssf);
        } else {
            XSSFSheet sheetXssf = (XSSFSheet) sheet;
            return getSheetPictures07(sheetNum, sheetXssf);
        }
    }

    /**
     * 获取Excel2003图片
     *
     * @param sheetNum 当前sheet编号
     * @param sheet    当前sheet对象
     * @return Map key:图片单元格索引（0-sheet下标,1-列号,1-行号）String，value:图片流PictureData
     */
    private static Map<String, PictureData> getSheetPictures03(int sheetNum, HSSFSheet sheet) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<>();
        List<HSSFPictureData> pictures = sheet.getWorkbook().getAllPictures();
        if (!pictures.isEmpty()) {
            HSSFPatriarch drawingPatriarch = sheet.getDrawingPatriarch();
            if (drawingPatriarch != null) {
                for (HSSFShape shape : drawingPatriarch.getChildren()) {
                    HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                    if (shape instanceof HSSFPicture) {
                        HSSFPicture pic = (HSSFPicture) shape;
                        int pictureIndex = pic.getPictureIndex() - 1;
                        HSSFPictureData picData = pictures.get(pictureIndex);
                        String picIndex = sheetNum + "," + anchor.getRow1() + "," + anchor.getCol1();
                        sheetIndexPicMap.put(picIndex, picData);
                    }
                }
            }
        }
        return sheetIndexPicMap;
    }

    /**
     * 获取Excel2007图片
     *
     * @param sheetNum 当前sheet编号
     * @param sheet    当前sheet对象
     * @return Map key:图片单元格索引（0,1,1）String，value:图片流PictureData
     */
    private static Map<String, PictureData> getSheetPictures07(int sheetNum, XSSFSheet sheet) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<>();
        for (POIXMLDocumentPart dr : sheet.getRelations()) {
            if (dr instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) dr;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    if (shape instanceof XSSFPicture) {
                        XSSFPicture pic = (XSSFPicture) shape;
                        XSSFClientAnchor anchor = pic.getPreferredSize();
                        CTMarker ctMarker = anchor.getFrom();
                        String picIndex = sheetNum + "," + ctMarker.getRow() + "," + ctMarker.getCol();
                        sheetIndexPicMap.put(picIndex, pic.getPictureData());
                    }
                }
            }
        }
        return sheetIndexPicMap;
    }

    /**
     * excel添加下拉数据校验
     *
     * @param sheet      哪个 sheet 页添加校验
     * @param dataSource 数据源数组
     * @param col        第几列校验（0开始）
     * @param firstRow   开始行
     * @param lastRow    结束行
     * @return DataValidation
     */
    private static DataValidation createDropDownValidation(Sheet sheet, String[] dataSource, int col, int firstRow, int lastRow) {
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(firstRow, lastRow, col, col);
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint;
        if (sheet.getWorkbook() instanceof HSSFWorkbook) {
            constraint = helper.createExplicitListConstraint(dataSource);
        } else {
            Workbook workbook = sheet.getWorkbook();
            Sheet hidden = workbook.getSheet("hidden");
            if (hidden == null) {
                hidden = workbook.createSheet("hidden");
            }
            // 1.首先创建行
            int dataLength = dataSource.length;
            int rowNum = hidden.getLastRowNum();
            char colLetter = 'A';
            if (rowNum == -1) {
                // 第一次创建下拉框数据
                for (int i = 0; i < dataLength; i++, rowNum++) {
                    hidden.createRow(i).createCell(0).setCellValue(dataSource[i]);
                }
            } else {
                // 之前已经创建过
                int createNum = dataLength - ++rowNum;
                short lastCellNum = hidden.getRow(0).getLastCellNum();
                for (int i = 0; i < lastCellNum; i++) {
                    colLetter++;
                }
                for (int i = 0; i < rowNum + createNum; i++) {
                    Row row = hidden.getRow(i);
                    if (row == null) {
                        row = hidden.createRow(i);
                    }
                    row.createCell(lastCellNum).setCellValue(dataSource[i]);
                }
            }
            // 3.设置表达式
            String formula = "hidden!$" + colLetter + "$1:$" + colLetter + "$" + dataLength;
            constraint = helper.createFormulaListConstraint(formula);
            workbook.setSheetHidden(1, true);
        }
        DataValidation dataValidation = helper.createValidation(constraint, cellRangeAddressList);

        // 处理Excel兼容性问题
        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        dataValidation.setEmptyCellAllowed(true);
        dataValidation.setShowPromptBox(true);
        dataValidation.createPromptBox("提示", "只能选择下拉框里面的数据");
        return dataValidation;
    }

    /**
     * excel添加时间数据校验
     *
     * @param sheet  哪个 sheet 页添加校验
     * @param start  開始
     * @param end    结束
     * @param info   提示信息
     * @param col    第几列校验（0开始）
     * @param maxRow 表头占用几行
     * @return DataValidation
     */
    private static DataValidation createDateValidation(Sheet sheet, String pattern, String start, String end, String info, int col, int maxRow, int lastRow) throws Exception {
        // 1.设置验证
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(maxRow, lastRow, col, col);
        DataValidationHelper helper = sheet.getDataValidationHelper();
        Calendar cal = Calendar.getInstance();
        Date startDate = DateUtils.parseDate(start, pattern);
        Date endDate = DateUtils.parseDate(end, pattern);
        cal.setTime(startDate);
        String formulaStart = "=DATE(" + cal.get(Calendar.YEAR) + "," + (cal.get(Calendar.MONTH) + 1) + "," + cal.get(Calendar.DATE) + ")";
        cal.setTime(endDate);
        String formulaEnd = "=DATE(" + cal.get(Calendar.YEAR) + "," + (cal.get(Calendar.MONTH) + 1) + "," + cal.get(Calendar.DATE) + ")";
        DataValidationConstraint constraint = helper.createDateConstraint(OperatorType.BETWEEN, formulaStart, formulaEnd, pattern);
        DataValidation dataValidation = handleMultiVersion(info, cellRangeAddressList, helper, constraint);
        // 2.设置单元格格式
        Workbook workbook = sheet.getWorkbook();
        CellStyle style = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        style.setDataFormat(createHelper.createDataFormat().getFormat(pattern));
        sheet.setDefaultColumnStyle(col, style);
        return dataValidation;
    }

    /**
     * 兼容性问题处理
     *
     * @param info                 提示消息
     * @param cellRangeAddressList 地址
     * @param helper               验证器
     * @param constraint           验证
     * @return DataValidation
     */
    private static DataValidation handleMultiVersion(String info, CellRangeAddressList cellRangeAddressList, DataValidationHelper helper, DataValidationConstraint constraint) {
        DataValidation dataValidation = helper.createValidation(constraint, cellRangeAddressList);
        // 处理Excel兼容性问题
        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        dataValidation.setEmptyCellAllowed(true);
        dataValidation.setShowPromptBox(true);
        if (info != null) {
            dataValidation.createPromptBox("提示", info);
        }
        return dataValidation;
    }

    /**
     * excel添加数字校验
     *
     * @param sheet  哪个 sheet 页添加校验
     * @param minNum 最小值
     * @param maxNum 最大值
     * @param info   提示信息
     * @param col    第几列校验（0开始）
     * @param maxRow 表头占用几行
     * @return DataValidation
     */
    private static DataValidation createNumValidation(Sheet sheet, String minNum, String maxNum, String info, int col, int maxRow, int lastRow) {
        // 1.设置验证
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(maxRow, lastRow, col, col);
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createIntegerConstraint(OperatorType.BETWEEN, minNum, maxNum);
        return handleMultiVersion(info, cellRangeAddressList, helper, constraint);
    }

    /**
     * excel添加数字校验
     *
     * @param sheet  哪个 sheet 页添加校验
     * @param minNum 最小值
     * @param maxNum 最大值
     * @param col    第几列校验（0开始）
     * @param maxRow 表头占用几行
     * @return DataValidation
     */
    private static DataValidation createFloatValidation(Sheet sheet, String minNum, String maxNum, String info, int col, int maxRow, int lastRow) {
        // 1.设置验证
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(maxRow, lastRow, col, col);
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createDecimalConstraint(OperatorType.BETWEEN, minNum, maxNum);
        return handleMultiVersion(info, cellRangeAddressList, helper, constraint);
    }

    /**
     * excel添加文本字符长度校验
     *
     * @param sheet  哪个 sheet 页添加校验
     * @param minNum 最小值
     * @param maxNum 最大值
     * @param info   自定义提示
     * @param col    第几列校验（0开始）
     * @param maxRow 表头占用几行
     * @return DataValidation
     */
    private static DataValidation createTextLengthValidation(Sheet sheet, String minNum, String maxNum, String info, int col, int maxRow, int lastRow) {
        // 1.设置验证
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(maxRow, lastRow, col, col);
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createTextLengthConstraint(OperatorType.BETWEEN, minNum, maxNum);
        return handleMultiVersion(info, cellRangeAddressList, helper, constraint);
    }

    /**
     * excel添加自定义校验
     *
     * @param sheet   哪个 sheet 页添加校验
     * @param formula 表达式
     * @param col     第几列校验（0开始）
     * @param maxRow  表头占用几行
     * @return DataValidation
     */
    private static DataValidation createCustomValidation(Sheet sheet, String formula, String info, int col, int maxRow, int lastRow) {
        String msg = "请输入正确的值！";
        // 0.修正xls表达式不正确定位的问题,只修正了开始，如F3:F2000,修正了F3变为A2,F2000变为A2000
        Workbook workbook = sheet.getWorkbook();
        if (workbook instanceof HSSFWorkbook) {
            int start = formula.indexOf("(") + 1;
            int end = formula.indexOf(")");
            if (start != 1 && end != 0) {
                String prev = formula.substring(0, start);
                String suffix = formula.substring(end);
                String substring = formula.substring(start, end);
                String[] ranges = substring.split(":");
                StringBuilder chars = new StringBuilder();
                Pattern pattern = Pattern.compile("([A-Z]+)(\\d+)");
                for (String range : ranges) {
                    Matcher matcher = pattern.matcher(range);
                    if (matcher.find()) {
                        int rowNum = Integer.parseInt(matcher.group(2));
                        chars.append("A").append(rowNum - 1).append(":");
                    }

                }
                chars.deleteCharAt(chars.length() - 1);
                formula = prev + chars + suffix;
            }

        }
        // 1.设置验证
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(maxRow, lastRow, col, col);
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createCustomConstraint(formula);
        DataValidation dataValidation = helper.createValidation(constraint, cellRangeAddressList);

        // 处理Excel兼容性问题
        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        dataValidation.setEmptyCellAllowed(true);
        dataValidation.setShowPromptBox(true);
        if (info != null) {
            msg = info;
        }
        dataValidation.createPromptBox("提示", msg);
        return dataValidation;
    }

}
