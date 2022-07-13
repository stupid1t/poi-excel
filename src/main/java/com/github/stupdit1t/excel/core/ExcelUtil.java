package com.github.stupdit1t.excel.core;

import com.github.stupdit1t.excel.callback.InCallback;
import com.github.stupdit1t.excel.common.*;
import com.github.stupdit1t.excel.core.export.ComplexCell;
import com.github.stupdit1t.excel.core.export.ExportRules;
import com.github.stupdit1t.excel.core.export.OutColumn;
import com.github.stupdit1t.excel.core.parse.InColumn;
import com.github.stupdit1t.excel.handle.ImgHandler;
import com.github.stupdit1t.excel.handle.rule.BaseVerifyRule;
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
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.security.GeneralSecurityException;
import java.text.ParseException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.function.BiConsumer;
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
	 * 设置打印方向
	 *
	 * @param sheet sheet页
	 */
	public static void printSetup(Sheet sheet) {
		PrintSetup printSetup = sheet.getPrintSetup();
		// 打印方向，true：横向，false：纵向
		printSetup.setLandscape(true);
		sheet.setFitToPage(true);
		sheet.setHorizontallyCenter(true);
	}

	/**
	 * 给工作簿加密码
	 *
	 * @param workbook 工作簿
	 * @param password 密码
	 */
	public static void encryptWorkbook03(Workbook workbook, String password) {
		// 2003
		Biff8EncryptionKey.setCurrentUserPassword(password);
		((HSSFWorkbook) workbook).writeProtectWorkbook(password, StringUtils.EMPTY);
	}

	/**
	 * 创建大数据workBook
	 * 避免OOM,导出速度比较慢
	 * <p>
	 * 默认后缀 xlsx
	 *
	 * @param rowAccessWindowSize 在内存中的行数
	 */
	public static Workbook createBigWorkbook(int rowAccessWindowSize) {
		return new SXSSFWorkbook(rowAccessWindowSize);
	}

	/**
	 * 创建空的workBook，做循环填充用
	 *
	 * @param xlsx 是否为xlsx格式
	 */
	public static Workbook createEmptyWorkbook(boolean xlsx) {
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
	public static void export(Workbook workbook, HttpServletResponse response, String fileName, String password) {
		export(workbook, getDownloadStream(response, fileName), password);
	}

	/**
	 * 导出
	 *
	 * @param out         导出流
	 * @param data        数据源
	 * @param exportRules 导出规则
	 */
	public static <T> void export(OutputStream out, List<T> data, ExportRules exportRules) {
		Workbook workbook = createEmptyWorkbook(exportRules.isXlsx());
		fillBook(workbook, data, exportRules);
		export(workbook, out, exportRules.getPassword());
	}


	/**
	 * 导出
	 *
	 * @param workbook 工作簿
	 * @param outPath  删除目录
	 */
	public static void export(Workbook workbook, String outPath, String password) {
		try (
				Workbook wb = workbook;
				OutputStream out = new FileOutputStream(outPath)
		) {
			export(wb, out, password);
		} catch (IOException e) {
			LOG.error(e);
		}
	}

    /**
     * 导出
     *
     * @param workbook     工作簿
     * @param outputStream 流
     */
    public static void export(Workbook workbook, OutputStream outputStream, String password) {
        try (
                Workbook wb = workbook;
                OutputStream out = outputStream
        ) {
            // 如果有密码, 且是03Excel
            if (StringUtils.isNotBlank(password)) {
                if (wb instanceof HSSFWorkbook) {
                    encryptWorkbook03(workbook, password);
                    wb.write(out);
                } else {
                    // 其它版本excel
                    EncryptionInfo info = new EncryptionInfo(EncryptionMode.standard);
                    Encryptor enc = info.getEncryptor();
                    enc.confirmPassword(password);
                    POIFSFileSystem poifsFileSystem = new POIFSFileSystem();
                    try {
                        OutputStream encOutStream = enc.getDataStream(poifsFileSystem);
                        wb.write(encOutStream);
                        encOutStream.close();
                        poifsFileSystem.writeFilesystem(out);
                        poifsFileSystem.close();
                    } catch (GeneralSecurityException e) {
                        LOG.error(e);
                    }
                }
            } else {
                wb.write(out);
            }
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
	public static <T> void fillBook(Workbook wb, List<T> data, ExportRules exportRules) {

		// -------------------- 全局样式处理 start ------------------------
		ICellStyle[] globalStyle = exportRules.getGlobalStyle();
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

		String sheetName = exportRules.getSheetName();
		Sheet sheet = safeCreateSheet(wb, sheetName);
		ExcelUtil.printSetup(sheet);

		// ----------------------- 表头设置 start ------------------------

		// 创建表头
		for (int i = 0; i < exportRules.getMaxRows(); i++) {
			Row row = sheet.createRow(i);
			for (int j = 0; j < exportRules.getMaxColumns(); j++) {
				row.createCell(j);
			}
		}
		// 合并模式
		if (exportRules.isIfMerge()) {
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
		Drawing<?> createDrawingPatriarch = safeCreateDrawing(sheet);
		// 存储类的字段信息
		Map<Class<?>, Map<String, Field>> clsInfo = new HashMap<>();
		// 存储单元格样式信息，防止重复生成
		Map<String, CellStyle> cacheStyle = new HashMap<>();
		// 存储单元格字体信息，防止重复生成
		Map<String, Font> cacheFont = new HashMap<>();
		// 列信息
		List<OutColumn<?>> fields = exportRules.getColumn();
		// 纵向合并信息计算存放
		Map<String, Integer[]> mergerRepeatCellsMap = new HashMap<>();
		List<Integer[]> mergerRepeatCells = new ArrayList<>();
		// 上一次行数据
		Map<String, String> lastRepeatKeyMap = new HashMap<>();
		for (int i = 0; i < data.size(); i++) {
			Row row = sheet.createRow(i + exportRules.getMaxRows());
			if (cellStyle.getHeight() != -1) {
				row.setHeight(cellStyle.getHeight());
			}
			// 行高自定义设置
			if (exportRules.getCellHeight() != -1) {
				row.setHeight(exportRules.getCellHeight());
			}
			T t = data.get(i);
			for (int j = 0, n = 0; n < fields.size(); j++, n++) {
				OutColumn<T> column = (OutColumn<T>) fields.get(n);
				Cell cell = row.createCell(j);
				cell.setCellStyle(cellStyleSource);
				// 1.序号设置
				if (exportRules.isAutoNum() && j == 0) {
					cell.setCellValue(i + 1);
					n--;
					continue;
				}
				// 2.读取Map/Object对应字段值
				if (clsInfo.get(t.getClass()) == null) {
					clsInfo.put(t.getClass(), PoiCommon.getAllFields(t.getClass()));
				}
				Object value = readField(clsInfo, t, column.getField());

				// 3.填充列值
				OutColumn.Style style = column.getStyle();
				if (column.getOutHandle() != null) {
					style = OutColumn.Style.clone(column.getStyle());
					value = column.getOutHandle().callback(value, t, style);
				}

				// 4.样式处理
				setCellStyle(wb, cellFont, cacheStyle, cacheFont, style, cell, value);

				// 5.设置单元格值
				setCellValue(createDrawingPatriarch, value, cell);

				// 6.批注添加
				String comment = column.getComment();
				if (StringUtils.isNotBlank(comment)) {
					// 表示需要用户添加批注
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
					Comment cellComment = createDrawingPatriarch.createCellComment(clientAnchor);
					cellComment.setAddress(cell.getAddress());
					cellComment.setString(richTextString);
					cell.setCellComment(cellComment);
				}

				// 7. 纵向合并判断
				if (column.getMergerRepeatFieldValue() != null) {
					StringBuilder repeatValue = new StringBuilder();
					if (column.getMergerRepeatFieldValue().length == 1 && column.getMergerRepeatFieldValue()[0].equals(column.getField())){
						repeatValue.append(value);
					}else{
						for (String repeatField : column.getMergerRepeatFieldValue()) {
							repeatValue.append(readField(clsInfo, t, repeatField));
						}
					}
					String nowKey = column.getField() + repeatValue;
					String lastRepeatKey = lastRepeatKeyMap.getOrDefault(column.getField(), "");
					if (!nowKey.equals(lastRepeatKey)) {
						// 内容不相同, 则重置上一次单元格数据, 存放合并数据
						Integer[] mergerRepeatCell = mergerRepeatCellsMap.remove(lastRepeatKey);
						if (mergerRepeatCell != null) {
							mergerRepeatCells.add(mergerRepeatCell);
						}
						lastRepeatKey = nowKey;
						lastRepeatKeyMap.put(column.getField(), lastRepeatKey);
					}
					Integer[] mergerCell = mergerRepeatCellsMap.getOrDefault(lastRepeatKey, new Integer[4]);
					mergerCell[2] = j;
					mergerCell[3] = j;
					if (mergerCell[0] == null) {
						mergerCell[0] = i + exportRules.getMaxRows();
						mergerCell[1] = i + exportRules.getMaxRows();
					} else {
						mergerCell[1] = i + exportRules.getMaxRows();
					}

					// 如果是最后一行, 则直接存放合并数据
					if (i == data.size() - 1) {
						mergerRepeatCells.add(mergerCell);
					} else {
						mergerRepeatCellsMap.put(lastRepeatKey, mergerCell);
					}
				}

			}
		}
		// ----------------------- body设置 end     -------------------------

		// ------------------------footer设置 start  ------------------------
		handleFooter(data, exportRules, footerFont, footerStyleSource, footerStyle, sheet);
		// ------------------------footer设置 end ---------------------------

		// ------------------------ 设置重复行合并 start ----------------------
		for (Integer[] mergerCell : mergerRepeatCells) {
			if (!mergerCell[0].equals(mergerCell[1])) {
				cellMerge(sheet, mergerCell[0], mergerCell[1], mergerCell[2], mergerCell[3]);
			}
		}
		// ------------------------ 设置重复行合并 end ------------------------

		// ------------------------ 设置自定义合并 start ----------------------
		List<Integer[]> mergerCells = exportRules.getMergerCells();
		if (mergerCells != null) {
			for (Integer[] mergerCell : mergerCells) {
				cellMerge(sheet, mergerCell[0], mergerCell[1], mergerCell[2], mergerCell[3]);
			}
		}

		// ------------------------ 设置自定义合并 end ------------------------
	}

	/**
	 * 同步创建drawing
	 *
	 * @param sheet sheet
	 * @return Drawing
	 */
	private static synchronized Drawing<?> safeCreateDrawing(Sheet sheet) {
		Drawing<?> createDrawingPatriarch;
		synchronized (ExcelUtil.class) {
			createDrawingPatriarch = sheet.createDrawingPatriarch();
		}
		return createDrawingPatriarch;
	}

	/**
	 * 同步创建sheet
	 *
	 * @param wb        工作簿
	 * @param sheetName sheet名字
	 * @return Sheet
	 */
	private static synchronized Sheet safeCreateSheet(Workbook wb, String sheetName) {
		Sheet sheet = sheetName != null ? wb.createSheet(sheetName) : wb.createSheet();
		return sheet;
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
	 */
	private static void setCellStyle(Workbook wb, Font cellFont, Map<String, CellStyle> cacheStyle, Map<String, Font> cacheFont, OutColumn.Style styleCustom, Cell cell, Object value) {
		String styleCacheKey = styleCustom.getStyleCacheKey();
		// 此处有值, 表示用户自定义列样式
		if (styleCacheKey != null) {
			CellStyle style = cacheStyle.get(styleCacheKey);
			// 表示缓存无, 重新构建
			if (style == null) {
				style = wb.createCellStyle();
				style.cloneStyleFrom(cell.getCellStyle());

				// 1.水平定位
				HorizontalAlignment align = styleCustom.getAlign();
				if (align != null) {
					style.setAlignment(align);
				}

				// 2.垂直定位
				VerticalAlignment valign = styleCustom.getValign();
				if (valign != null) {
					style.setVerticalAlignment(valign);
				}

				// 3.字体颜色
				IndexedColors color = styleCustom.getColor();
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
				IndexedColors backColor = styleCustom.getBackColor();
				if (backColor != null) {
					style.setFillForegroundColor(backColor.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				}

				// 5. 格式化
				String pattern = styleCustom.getPattern();
				if (StringUtils.isNotBlank(pattern)) {
					CreationHelper createHelper = wb.getCreationHelper();
					style.setDataFormat(createHelper.createDataFormat().getFormat(pattern));
				}

                // 6. 换行显示
                Boolean wrapText = styleCustom.getWrapText();
                if (wrapText != null) {
                    style.setWrapText(wrapText);
                }
				cacheStyle.put(styleCacheKey, style);
			}
			// 最终样式设置
			cell.setCellStyle(style);
		}

		// 如果是日期, 且用户没有设置日期格式化, 默认年月日时分秒
		boolean dateValue = value instanceof Date || value instanceof LocalDate || value instanceof LocalDateTime;
		if (dateValue && styleCustom.getPattern() == null) {
			String cacheDateKey = "global-signal-date";
			CellStyle style = cacheStyle.get(cacheDateKey);
			if (style == null) {
				style = wb.createCellStyle();
				style.cloneStyleFrom(cell.getCellStyle());
				CreationHelper createHelper = wb.getCreationHelper();
				style.setDataFormat(createHelper.createDataFormat().getFormat(PoiConstant.FMT_DATE_TIME));
				cacheStyle.put(cacheDateKey, style);
			}
			cell.setCellStyle(style);
		}

		// 6.高度
		int height = styleCustom.getHeight();
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
		if (exportRules.isIfFooter()) {
			Workbook workbook = sheet.getWorkbook();
			List<ComplexCell> footerRules = exportRules.getFooterRules();
			// 构建尾行
			int currRowNum = exportRules.getMaxRows() + data.size();
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
					if (exportRules.getFooterHeight() != -1) {
						sheet.getRow(i).setHeight(exportRules.getFooterHeight());
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
        List<OutColumn<?>> fields = exportRules.getColumn();
        int columnWidth = exportRules.getColumnWidth();
        int autoNumColumnWidth = exportRules.getAutoNumColumnWidth();
        for (int i = 0, j = 0; i < fields.size(); i++, j++) {
            // 0.每一列默认单元格样式设置
            // 1.width设置
            if (exportRules.isAutoNum() && j == 0) {
                j++;
                sheet.setColumnWidth(0, autoNumColumnWidth);
            }
            OutColumn<?> column = fields.get(i);
            // 1.1是否自动列宽
            int width = column.getStyle().getWidth();
            if (width != -1) {
                sheet.setColumnWidth(j, width);
            } else {
                try {
                    // 1.2根据maxRows，获取表头的值设置宽度
                    if (columnWidth != -1) {
                        sheet.setColumnWidth(j, columnWidth);
                    } else {
                        Row row = sheet.getRow(exportRules.getMaxRows() - 1);
                        String headerValue = row.getCell(j).getStringCellValue();
                        if (StringUtils.isBlank(headerValue)) {
                            row = sheet.getRow(exportRules.getMaxRows() - 2);
                            headerValue = row.getCell(j).getStringCellValue();
                        }
                        sheet.setColumnWidth(j, headerValue.getBytes().length * 256);
                    }
                } catch (Exception e) {
                    if (exportRules.isAutoNum()) {
                        LOG.error("请确认表头数量和列数量一致! ");
                        throw new UnsupportedOperationException("自动序号设置错误，请确认在header添加序号列");
                    } else {
                        LOG.error("请确认表头数量和列数量一致! ");
                        throw e;
                    }
                }
            }
            // 2.downDown设置
            int lastRow = (exportRules.getMaxRows() - 1) + data.size();
            lastRow = lastRow == (exportRules.getMaxRows() - 1) ? PoiConstant.MAX_FILL_COL : lastRow;
            String[] dropdown = column.getDropdown();
            if (dropdown != null && dropdown.length > 0) {
                sheet.addValidationData(createDropDownValidation(sheet, dropdown, j, exportRules.getMaxRows(), lastRow));
            }

			// 3.时间校验
			String date = column.getVerifyDate();
			if (date != null) {
				String[] split = date.split("@");
				String info = null;
				if (split.length == 2) {
					info = split[1];
				}
				String[] split1 = split[0].split("~");
				if (split1.length < 2) {
					throw new IllegalArgumentException("时间校验表达式不正确,请填写如" + column.getStyle().getPattern() + "的值!");
				}
				try {
					sheet.addValidationData(createDateValidation(sheet, column.getStyle().getPattern(), split1[0], split1[1], info, j, exportRules.getMaxRows(), lastRow));
				} catch (ParseException e) {
					LOG.error(e);
					throw new IllegalArgumentException("时间校验表达式不正确,请填写如" + column.getStyle().getPattern() + "的值!");
				} catch (Exception e) {
					LOG.error(e);
				}
			}

			// 4.整数数字校验
			String num = column.getVerifyIntNum();
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
				sheet.addValidationData(createNumValidation(sheet, split1[0], split1[1], info, j, exportRules.getMaxRows(), lastRow));
			}

			// 4.浮点数字校验
			String floatNum = column.getVerifyFloatNum();
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
				sheet.addValidationData(createFloatValidation(sheet, split1[0], split1[1], info, j, exportRules.getMaxRows(), lastRow));
			}

			// 5.自定义校验
			String custom = column.getVerifyCustom();
			if (custom != null) {
				String[] split = custom.split("@");
				String info = null;
				if (split.length == 2) {
					info = split[1];
				}
				sheet.addValidationData(createCustomValidation(sheet, split[0], info, j, exportRules.getMaxRows(), lastRow));
			}

			// 6.文本长度校验
			String text = column.getVerifyText();
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
				sheet.addValidationData(createTextLengthValidation(sheet, split2[0], split2[1], info, j, exportRules.getMaxRows(), lastRow));
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
		if (exportRules.isFreezeHeader()) {
			sheet.createFreezePane(0, exportRules.getMaxRows(), 0, exportRules.getMaxRows());
		}

		// 2. title 内容设置和行高
		if (exportRules.getTitle() != null) {
			// title全局行高
			if (titleStyle.getHeight() != -1) {
				sheet.getRow(0).setHeight(titleStyle.getHeight());
			}
			// title行高自定义设置
			if (exportRules.getTitleHeight() != -1) {
				sheet.getRow(0).setHeight(exportRules.getTitleHeight());
			}
			CellUtil.createCell(sheet.getRow(0), 0, exportRules.getTitle(), titleStyleSource);
			cellMerge(sheet, 0, 0, 0, exportRules.getMaxColumns());
		}

		// 3.header设置和行高
		int headerIndex = exportRules.getTitle() == null ? 0 : 1;
		// header全局行高
		if (headerStyle.getHeight() != -1) {
			sheet.getRow(headerIndex).setHeight(headerStyle.getHeight());
		}
		// header行高自定义设置
		if (exportRules.getHeaderHeight() != -1) {
			sheet.getRow(headerIndex).setHeight(exportRules.getHeaderHeight());
		}
		LinkedHashMap<String, BiConsumer<Font, CellStyle>> headerMap = exportRules.getSimpleHeader();
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
		if (exportRules.isFreezeHeader()) {
			sheet.createFreezePane(0, exportRules.getMaxRows(), 0, exportRules.getMaxRows());
		}
		// header
		List<ComplexCell> complexHeader = exportRules.getComplexHeader();
		for (ComplexCell complexCell : complexHeader) {
			Integer[] range = complexCell.getLocationIndex();
			// 合并表头
			int firstRow = range[0];
			int lastRow = range[1];
			int firstCol = range[2];
			int lastCol = range[3];
			CellStyle styleTemp;
			Font fontTemp;
			if ((exportRules.getMaxColumns() - 1) == lastCol - firstCol && firstRow == 0) {
				// 占满全格, 且第一行开始为表头
				for (int i = firstRow; i <= lastRow; i++) {
					if (titleStyle.getHeight() != -1) {
						sheet.getRow(i).setHeight(titleStyle.getHeight());
					}
					// 行高自定义设置
					if (exportRules.getTitleHeight() != -1) {
						sheet.getRow(i).setHeight(exportRules.getTitleHeight());
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
					if (exportRules.getHeaderHeight() != -1) {
						sheet.getRow(i).setHeight(exportRules.getHeaderHeight());
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
	 * 读取规则excel数据内容为map
	 *
	 * @param filePath     文件路径
	 * @param poiSheetArea 数据区域
	 * @param columns      数据列定义
	 * @param callBack     回调数据行
	 * @param rowClass     数据类
	 * @return PoiResult
	 */
	public static <T> PoiResult<T> readSheet(String filePath, PoiSheetDataArea poiSheetArea, Map<String, InColumn<?>> columns, InCallback<T> callBack, Class<T> rowClass) {
		try (InputStream is = new FileInputStream(filePath)) {
			return readSheet(is, poiSheetArea, columns, callBack, rowClass);
		} catch (IOException e) {
			LOG.error(e);
		}
		return new PoiResult<>();
	}

	/**
	 * 读取规则excel数据内容为map
	 *
	 * @param is           文件流
	 * @param poiSheetArea 数据区域
	 * @param columns      数据列定义
	 * @param callBack     回调数据行
	 * @param rowClass     数据类
	 * @return PoiResult
	 */
	public static <T> PoiResult<T> readSheet(InputStream is, PoiSheetDataArea poiSheetArea, Map<String, InColumn<?>> columns, InCallback<T> callBack, Class<T> rowClass) {
		try (Workbook wb = WorkbookFactory.create(is)) {
			String sheetName = poiSheetArea.getSheetName();
			Sheet sheet;
			if (StringUtils.isBlank(sheetName)) {
				sheet = wb.getSheetAt(poiSheetArea.getSheetIndex());
			} else {
				sheet = wb.getSheet(sheetName);
			}
			return readSheet(sheet, poiSheetArea.getHeaderRowCount(), poiSheetArea.getFooterRowCount(), columns, callBack, rowClass);
		} catch (Exception e) {
			LOG.error(e);
		}
		return new PoiResult<>();
	}

	/**
	 * 读取规则excel数据内容为map
	 *
	 * @param sheet           sheet页
	 * @param dataStartRow    起始行
	 * @param dataEndRowCount 尾部非数据行数量
	 * @return PoiResult<T>
	 */
	public static <T> PoiResult<T> readSheet(Sheet sheet, int dataStartRow, int dataEndRowCount, Map<String, InColumn<?>> columns, InCallback<T> callBack, Class<T> rowClass) {
		boolean mapClass = PoiCommon.isMapData(rowClass);
		PoiResult<T> rsp = new PoiResult<>();
		List<T> beans = new ArrayList<>();
		// 获取excel中所有图片
		Set<String> hasImgField = new HashSet<>();
		Map<String, PictureData> pictures = null;
		Collection<InColumn<?>> values = columns.values();
		int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
		for (InColumn<?> inColumn : values) {
			BaseVerifyRule<?> cellVerify = inColumn.getCellVerifyRule();
			if (cellVerify instanceof ImgHandler) {
				if (pictures == null) {
					pictures = getSheetPictures(sheetIndex, sheet);
				}
				hasImgField.add(inColumn.getField());
			}
		}

		// 公式计算初始化
		FormulaEvaluator formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		int rowStart = sheet.getFirstRowNum() + dataStartRow;
		// 获取真实的数据行尾数
		int rowEnd = getLastRealLastRow(sheet.getRow(sheet.getLastRowNum())) - dataEndRowCount;
		List<String> errors = new ArrayList<>();
		try {
			for (int j = rowStart; j <= rowEnd; j++) {
				List<String> rowErrors = new ArrayList<>();
				T data = rowClass.newInstance();
				Row row = sheet.getRow(j);
				if (row == null) {
					continue;
				}
				int lastCellNum = columns.size() == 0 ? row.getLastCellNum() : columns.size();
				for (int k = 0; k < lastCellNum; k++) {
					String fieldName;
					try {
						// 列名称获取
						String columnIndexChar = PoiConstant.numsRefCell.get(k);
						InColumn<?> inColumn = columns.get(columnIndexChar);
						Object cellValue;
						if (inColumn != null) {
							fieldName = inColumn.getField();
						} else {
							fieldName = columnIndexChar;
						}

						if (pictures != null && hasImgField.contains(fieldName)) {
							String pictureIndex = sheetIndex + "," + j + "," + k;
							PictureData remove = pictures.remove(pictureIndex);
							cellValue = remove == null ? null : remove.getData();
						} else {
							cellValue = getCellValue(row, k, formulaEvaluator);
						}

						// 校验类型转换处理
						if (inColumn != null) {
							cellValue = inColumn.getCellVerifyRule().handle(inColumn.getTitle(), columnIndexChar + (j + 1), cellValue);
						}

						if (mapClass) {
							((Map) data).put(fieldName, cellValue);
						} else {
							FieldUtils.writeField(data, fieldName, cellValue, true);
						}
					} catch (PoiException e) {
						rowErrors.add(e.getMessage());
					}
				}
				// 如果行错误不为空, 添加错误
				if (!rowErrors.isEmpty()) {
					errors.add(String.format(PoiConstant.ROW_INDEX_STR, j + 1, String.join(" ", rowErrors)));
				} else {
					// 有效, 回调处理加入
					if (callBack != null) {
						callBack.callback(data, j + 1);
					}
					beans.add(data);
				}
			}
		} catch (Exception e) {
			LOG.error(e);
		} finally {
			// throw parse exception
			if (errors.size() > 0) {
				rsp.setSuccess(false);
				rsp.setMessage(errors);
			}
			rsp.setData(beans);
		}
		// 返回结果
		return rsp;
	}

	/**
	 * 读取excel,替换内置变量
	 *
	 * @param filePath 文件路径
	 * @param variable 内置变量
	 */
	public static Workbook readExcelWrite(String filePath, Map<String, String> variable) {
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
	public static Workbook readExcelWrite(InputStream is, Map<String, String> variable) {
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
	private static void setCellValue(Drawing<?> createDrawingPatriarch, Object value, Cell cell) {
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
			cell.setCellValue(((Number) value).doubleValue());
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
				anchor = new XSSFClientAnchor(0, 0, 0, 0, x, y, x + 1, y + 1);
				add1 = workbook.addPicture(data, XSSFWorkbook.PICTURE_TYPE_PNG);
			} else if (workbook instanceof HSSFWorkbook) {
				anchor = new HSSFClientAnchor(0, 0, 0, 0, x, y, (short) (x + 1), y + 1);
				add1 = workbook.addPicture(data, SXSSFWorkbook.PICTURE_TYPE_PNG);
			} else {
				anchor = new XSSFClientAnchor(0, 0, 0, 0, x, y, (short) (x + 1), y + 1);
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
				// 拿到计算公式eval, 捕捉公式错误异常
				try {
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
				} catch (Exception e) {
					obj = "";
					LOG.error("公式有误:{0}", e);
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
		if (pictures.isEmpty()) {
			return sheetIndexPicMap;
		}
		HSSFPatriarch drawingPatriarch = sheet.getDrawingPatriarch();
		if (drawingPatriarch == null) {
			return sheetIndexPicMap;
		}
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
						XSSFClientAnchor anchor = pic.getClientAnchor();
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
