package excel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import excel.callBack.ExportSheetCallback;
import excel.callBack.ParseSheetCallback;
import excel.verify.AbstractVerifyBuidler;

public class ExcelUtils {
	private static final Logger LOG = LoggerFactory.getLogger(ExcelUtils.class);

	/**
	 * 设置打印方向
	 * 
	 * @param sheet
	 */
	private static void printSetup(Sheet sheet) {
		PrintSetup printSetup = sheet.getPrintSetup();
		// 打印方向，true：横向，false：纵向
		printSetup.setLandscape(true);
		sheet.setFitToPage(true);
		sheet.setHorizontallyCenter(true);
	}

	/**
	 * 内置初始化样式
	 * 
	 * @param wb
	 * @return
	 */
	private static Map<String, CellStyle> initStyles(Workbook wb) {
		Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
		CellStyle style;
		Font titleFont = wb.createFont();
		titleFont.setFontHeightInPoints((short) 15);
		titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		style = wb.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 左右居中
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 上下居中
		style.setFont(titleFont);
		styles.put("title", style);

		Font monthFont = wb.createFont();
		monthFont.setFontName("Arial");
		monthFont.setFontHeightInPoints((short) 10);
		monthFont.setColor(IndexedColors.WHITE.getIndex());
		style = wb.createCellStyle();
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setFont(monthFont);
		style.setWrapText(true);
		styles.put("header", style);

		style = wb.createCellStyle();
		Font cellFont = wb.createFont();
		cellFont.setFontName("Arial");
		cellFont.setFontHeightInPoints((short) 10);
		style.setFont(cellFont);
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setWrapText(false);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		styles.put(POIConstant.CENTER, style);

		style = wb.createCellStyle();
		style.setFont(cellFont);
		style.setAlignment(CellStyle.ALIGN_LEFT);
		style.setWrapText(false);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		styles.put(POIConstant.LEFT, style);

		style = wb.createCellStyle();
		style.setFont(cellFont);
		style.setAlignment(CellStyle.ALIGN_RIGHT);
		style.setWrapText(false);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		styles.put(POIConstant.RIGHT, style);

		return styles;
	}

	/**
	 * @param list         数据源
	 * @param hearderRules （如带序号，在规则里设置序头） 表头规则
	 * @param autoNum      带序号
	 * @param fields       字段，长度
	 * @return
	 * @throws Exception
	 */
	private static <T> Workbook createWorkbook(List<T> list, boolean autoNum, ExportRules hearderRules, Object[][] fields, ExportSheetCallback<T> callBack, boolean is07) throws Exception {
		Workbook wb = null;
		if (is07) {
			wb = new XSSFWorkbook();// 2007
		} else {
			wb = new HSSFWorkbook();// 2003
		}
		Map<String, CellStyle> styles = ExcelUtils.initStyles(wb);
		Sheet sheet = wb.createSheet();
		ExcelUtils.printSetup(sheet);
		int maxColumns = hearderRules.getMaxColumns();
		int maxRows = hearderRules.getMaxRows();
		if (hearderRules.getIfMerge()) {// 合并模式
			// 冻结表头
			sheet.createFreezePane(0, maxRows, 0, maxRows);
			// header
			HashMap<String, String> rules = hearderRules.getHeaderRules();
			for (int i = 0; i < maxRows; i++) {
				sheet.createRow(i);
			}
			Iterator<Entry<String, String>> entries = rules.entrySet().iterator();
			while (entries.hasNext()) {
				Entry<String, String> entry = entries.next();
				String key = entry.getKey();
				String value = entry.getValue();
				Object[] range = coverRange(key);
				// 合并表头
				CellRangeAddress cra = new CellRangeAddress((int) range[0] - 1, (int) range[1] - 1, POIConstant.cellRefNums.get(range[2]), POIConstant.cellRefNums.get(range[3]));
				sheet.addMergedRegion(cra);
				RegionUtil.setBorderBottom(1, cra, sheet, wb); // 下边框
				RegionUtil.setBorderLeft(1, cra, sheet, wb); // 左边框
				RegionUtil.setBorderRight(1, cra, sheet, wb); // 有边框
				RegionUtil.setBorderTop(1, cra, sheet, wb); // 上边框
				if ((maxColumns - 1) == POIConstant.cellRefNums.get(range[3]) - POIConstant.cellRefNums.get(range[2])) {// 占满全格，则为表头
					CellUtil.createCell(sheet.getRow((int) range[0] - 1), POIConstant.cellRefNums.get(range[2]), value, styles.get("title"));
				} else {
					CellUtil.createCell(sheet.getRow((int) range[0] - 1), POIConstant.cellRefNums.get(range[2]), value, styles.get("header"));
				}
			}
		} else {// 非合并
			if (hearderRules.getTitile() == null) {
				// 冻结表头
				sheet.createFreezePane(0, 1, 0, 1);
				sheet.createRow(0);
				String[] hearder = hearderRules.getHearder();
				for (int i = 0; i < hearder.length; i++) {
					CellUtil.createCell(sheet.getRow(0), i, hearder[i], styles.get("header"));
				}
			} else {
				// 冻结表头
				sheet.createFreezePane(0, 2, 0, 2);
				sheet.createRow(0);
				sheet.createRow(1);
				CellRangeAddress cra = new CellRangeAddress(0, 0, 0, maxColumns);
				sheet.addMergedRegion(cra);
				CellUtil.createCell(sheet.getRow(0), 0, hearderRules.getTitile(), styles.get("title"));
				String[] hearder = hearderRules.getHearder();
				for (int i = 0; i < hearder.length; i++) {
					CellUtil.createCell(sheet.getRow(1), i, hearder[i], styles.get("header"));
				}
			}

		}
		// set width
		if (autoNum) {
			sheet.setColumnWidth(0, 2000);
		}
		for (int i = 0, j = 0; i < fields.length; i++, j++) {
			if (autoNum) {
				j = i + 1;
			}
			Object[] columnsWidth = fields[i];
			// 是否自动列宽
			int width = (int) columnsWidth[1];
			if (width != 0) {
				sheet.setColumnWidth(j, width);
			} else {
				// 根据maxRows，获取表头的值设置宽度
				Row row = sheet.getRow(maxRows - 1);
				String headerValue = row.getCell(j).getStringCellValue();
				if (StringUtils.isBlank(headerValue)) {
					row = sheet.getRow(maxRows - 2);
					headerValue = row.getCell(j).getStringCellValue();
				}
				sheet.setColumnWidth(j, headerValue.getBytes().length * 256);
			}
		}
		// body row
		Drawing createDrawingPatriarch = sheet.createDrawingPatriarch();
		// 存储类的字段信息
		Map<Class<? extends Object>, Map<String, Field>> clsInfo = new HashMap<>();
		for (int i = 0; i < list.size(); i++) {
			Row row = sheet.createRow(i + maxRows);
			T t = list.get(i);
			for (int j = 0, n = 0; n < fields.length; j++, n++) {
				Cell cell = row.createCell(j);

				// 有效数据行号
				if (autoNum && j == 0) {
					cell.setCellStyle(styles.get(POIConstant.CENTER));
					cell.setCellValue(i + 1);
					n--;
					continue;
				}

				// 读取Map/Object对应字段值
				if (clsInfo.get(t.getClass()) == null) {
					clsInfo.put(t.getClass(), getAllFields(t.getClass()));
				}
				Object value = readField(clsInfo, t, (String) fields[n][0]);

				// 填充列值
				CellStyle style = styles.get(POIConstant.CENTER);
				if (callBack != null) {
					value = callBack.callback((String) fields[n][0], value, t);
				}

				// 自定义样式
				if (fields[n].length == 3) {
					String styleType = (String) fields[n][2];
					if (POIConstant.LEFT.equals(styleType)) {
						style = styles.get(POIConstant.LEFT);
					} else if (POIConstant.RIGHT.equals(styleType)) {
						style = styles.get(POIConstant.RIGHT);
					}
				}
				// 设置单元格值
				if (value instanceof byte[]) {
					byte[] data = (byte[]) value;
					// anchor主要用于设置图片的属性
					short x = (short) cell.getColumnIndex();
					int y = cell.getRowIndex();
					// 插入图片
					XSSFClientAnchor anchor = new XSSFClientAnchor(10, 10, 10, 10, x, y, x + 1, y + 1);
					int add1 = wb.addPicture(data, XSSFWorkbook.PICTURE_TYPE_PNG);
					createDrawingPatriarch.createPicture(anchor, add1);
					cell.setCellValue("");
				} else {
					setCellValue(style, value, cell);
				}
			}
		}
		// footer row
		if (hearderRules.getIfFooter()) {
			HashMap<String, String> footerRules = hearderRules.getFooterRules();
			// 构建尾行数字
			int currRownum = hearderRules.getMaxRows() + list.size();
			int[] footerNum = getFooterNum(footerRules.entrySet().iterator(), currRownum);
			Iterator<Entry<String, String>> entries = footerRules.entrySet().iterator();
			for (int i = 0; i < footerNum.length; i++) {
				sheet.createRow(footerNum[i]);
			}
			while (entries.hasNext()) {
				Entry<String, String> entry = entries.next();
				String key = entry.getKey();
				String value = entry.getValue();
				Object[] range = coverRange(key);
				CellRangeAddress cra = new CellRangeAddress((int) range[0] + currRownum - 1, (int) range[1] + currRownum - 1, POIConstant.cellRefNums.get(range[2]),
						POIConstant.cellRefNums.get(range[3]));
				sheet.addMergedRegion(cra);
				RegionUtil.setBorderBottom(1, cra, sheet, wb); // 下边框
				RegionUtil.setBorderLeft(1, cra, sheet, wb); // 左边框
				RegionUtil.setBorderRight(1, cra, sheet, wb); // 有边框
				RegionUtil.setBorderTop(1, cra, sheet, wb); // 上边框
				String cellValue = "";
				CellStyle style = styles.get(POIConstant.CENTER);
				cellValue = value;
				CellUtil.createCell(sheet.getRow((int) range[0] + currRownum - 1), POIConstant.cellRefNums.get(range[2]), cellValue, style);
			}

		}
		return wb;
	}

	/**
	 * 读取字段的值
	 * 
	 * @param clsInfo 类信息
	 * @param t       当前值
	 * @param fields  字段名称
	 * @return
	 * @throws Exception
	 */
	private static Object readField(Map<Class<? extends Object>, Map<String, Field>> clsInfo, Object t, String fields) throws Exception {
		// 返回值
		Object value = null;
		try {
			// 读取子属性
			String[] split = fields.split("\\.");
			if (t instanceof Map) {
				for (int i = 0; i < split.length; i++) {
					if (i == 0) {
						value = ((Map) t).get(split[i]);
					} else {
						if (value instanceof Map) {
							value = ((Map) value).get(split[i]);
						} else {
							Class<? extends Object> subCls = value.getClass();
							if (clsInfo.get(subCls) == null) {
								Map<String, Field> subField = getAllFields(subCls);
								clsInfo.put(subCls, subField);
							}
							Field field = clsInfo.get(subCls).get(split[i]);
							if (field == null) {
								// 为方法，不是属性
								char[] charName = split[i].toCharArray();
								charName[0] -= 32;
								String methodName = "get" + String.valueOf(charName);
								Method method = subCls.getMethod(methodName);
								value = method.invoke(value);
							} else {
								value = field.get(value);
							}
						}
					}
					// 属性为空跳出
					if (value == null) {
						break;
					}
				}
			} else {
				for (int i = 0; i < split.length; i++) {
					if (i == 0) {
						Field field = clsInfo.get(t.getClass()).get(split[i]);
						if (field == null) {
							// 为方法，不是属性
							char[] charName = split[i].toCharArray();
							charName[0] -= 32;
							String methodName = "get" + String.valueOf(charName);
							Method method = t.getClass().getMethod(methodName);
							value = method.invoke(value);
						} else {
							value = field.get(t);
						}
					} else {
						if (value instanceof Map) {
							value = ((Map) value).get(split[i]);
						} else {
							Class<? extends Object> subCls = value.getClass();
							if (clsInfo.get(subCls) == null) {
								Map<String, Field> subField = getAllFields(subCls);
								clsInfo.put(subCls, subField);
							}
							Field field = clsInfo.get(subCls).get(split[i]);
							if (field == null) {
								// 为方法，不是属性
								char[] charName = split[i].toCharArray();
								charName[0] -= 32;
								String methodName = "get" + String.valueOf(charName);
								Method method = subCls.getMethod(methodName);
								value = method.invoke(value);
							} else {
								value = field.get(value);
							}
						}

					}
					// 属性为空跳出
					if (value == null) {
						break;
					}
				}
			}
		} catch (Exception e) {
			value = null;
		}
		return value == null ? "" : value;
	}

	/**
	 * 获取实体的所有字段
	 * 
	 * @param t
	 * @return
	 */
	private static Map<String, Field> getAllFields(Class<?> t) {
		Map<String, Field> field = new HashMap<>();
		List<Field> allFieldsList = FieldUtils.getAllFieldsList(t);
		allFieldsList.stream().forEach(n -> {
			n.setAccessible(true);
			field.put(n.getName(), n);
		});
		return field;
	}

	/**
	 * 给单元格设置值
	 * 
	 * @param value   列值
	 * @param pattern 格式化值
	 * @param cell    单元格
	 * @param wb
	 */
	private static void setCellValue(CellStyle style, Object value, Cell cell) {
		cell.setCellStyle(style);
		// 判断值的类型后进行强制类型转换
		if (value instanceof String) {
			cell.setCellValue(String.valueOf(value));
		} else if (value instanceof Integer) {
			cell.setCellValue((Integer) (value));
		} else if (value instanceof Double) {
			DecimalFormat fmt = new DecimalFormat("#0.00");
			cell.setCellValue(fmt.format((Double) (value)));
		} else if (value instanceof Long) {
			cell.setCellValue((Long) (value));
		} else if (value instanceof Date) {
			cell.setCellValue(new SimpleDateFormat(POIConstant.FMTDATETIME).format((Date) (value)));
		} else if (value == null) {
			cell.setCellValue("");
		} else {
			cell.setCellValue(String.valueOf(value));
		}
	}

	/**
	 * 合并单元格转换
	 * 
	 * @param obj
	 * @return
	 */
	private static Object[] coverRange(Object obj) {
		String[] range = ((String) obj).split(",");
		Object[] rangeInt = new Object[4];
		for (int i = 0; i < range.length; i++) {
			if (i < 2) {
				rangeInt[i] = Integer.parseInt(range[i]);
			} else {
				rangeInt[i] = range[i];
			}

		}
		return rangeInt;
	}

	/**
	 * 根据页脚数据获得行号
	 * 
	 * @param entries
	 * @param currRownum
	 * @return
	 */
	private static int[] getFooterNum(Iterator<Entry<String, String>> entries, int currRownum) {
		int row = 0;
		while (entries.hasNext()) {
			Entry<String, String> entry = entries.next();
			String key = entry.getKey();
			Object[] range = coverRange(key);
			int a = (int) range[1];
			row = a > row ? a : row;
		}
		int[] footerNum = new int[row];
		for (int i = 0; i < row; i++) {
			footerNum[i] = currRownum + i;
		}
		return footerNum;
	}

	/**
	 * 导出带回调处理的
	 * 
	 * @param             <T>
	 * 
	 * @param data        数据源
	 * @param exportRules 导出规则
	 * @param callBack    回调处理逻辑
	 * @param isXlsx      是否导出xlsx格式
	 * @return
	 */
	public static <T> Workbook createWorkbook(List<T> data, ExportRules exportRules, ExportSheetCallback<T> callBack, boolean... isXlsx) {
		boolean is07Excel = false;
		if (isXlsx.length > 0) {
			is07Excel = isXlsx[0];
		}
		Workbook work = null;
		try {
			work = createWorkbook(data, exportRules.getAutoNum(), exportRules, exportRules.getFields(), callBack, is07Excel);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return work;
	}

	/**
	 * 导出
	 * 
	 * @param             <T>
	 * 
	 * @param data        数据源
	 * @param exportRules 导出规则
	 * @param is07        是否导出xlsx格式
	 * @return
	 */
	public static <T> Workbook createWorkbook(List<T> data, ExportRules exportRules, boolean... isXlsx) {
		boolean is07Excel = false;
		if (isXlsx.length > 0) {
			is07Excel = isXlsx[0];
		}
		Workbook work = null;
		try {
			work = createWorkbook(data, exportRules.getAutoNum(), exportRules, exportRules.getFields(), null, is07Excel);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return work;
	}

	/**
	 * 解析Sheet
	 * 
	 * @param clss            结果bean
	 * @param verifyBuilder   校验器
	 * @param sheet           解析的sheet
	 * @param dataStartRow    开始行:从0开始计
	 * @param dataEndRowCount 尾行非数据行数量
	 * @return
	 * @throws Exception
	 */
	public static <T> ImportRspInfo<T> parseSheet(Class<T> clss, AbstractVerifyBuidler verifyBuilder, Sheet sheet, int dataStartRow, int dataEndRowCount) {
		return parseSheet(clss, verifyBuilder, sheet, dataStartRow, dataEndRowCount, null);
	}

	/**
	 * 解析Sheet
	 * 
	 * @param clss            结果bean
	 * @param verifyBuilder   校验器
	 * @param sheet           解析的sheet
	 * @param dataStartRow    开始行
	 * @param dataEndRowCount 尾行数量
	 * @param callback        加入回调逻辑
	 * @return
	 * @throws Exception
	 */
	public static <T> ImportRspInfo<T> parseSheet(Class<T> clss, AbstractVerifyBuidler verifyBuilder, Sheet sheet, int dataStartRow, int dataEndRowCount, ParseSheetCallback<T> callback) {
		ImportRspInfo<T> rsp = new ImportRspInfo<T>();
		List<T> beans = new ArrayList<>();
		StringBuffer errors = new StringBuffer();
		StringBuffer rowErrors = new StringBuffer();
		try {
			int rowStart = sheet.getFirstRowNum() + dataStartRow;
			int rowEnd = sheet.getLastRowNum() - dataEndRowCount;
			for (int rowNum = rowStart; rowNum <= rowEnd; rowNum++) {
				Row r = sheet.getRow(rowNum);
				// 创建对象
				T t = clss.newInstance();
				int fieldNum = 0;
				for (int cellNum : POIConstant.convertToCellNum(verifyBuilder.cellRefs)) {
					// 列坐标
					CellReference cellRef = new CellReference(rowNum, cellNum);
					try {
						Object cellValue = getCellValue(r, cellNum);
						// 校验和格式化列值
						cellValue = verifyBuilder.verify(verifyBuilder.filedNames[fieldNum], cellValue);
						// 填充列值
						FieldUtils.writeDeclaredField(t, verifyBuilder.filedNames[fieldNum], cellValue, true);
					} catch (POIException e) {
						rowErrors.append(cellRef.formatAsString()).append(":").append(e.getMessage()).append("\t");
						LOG.error(e.getMessage());
					}
					fieldNum++;
				}
				// 回调处理一下特殊逻辑
				if (callback != null) {
					callback.callback(t, rowNum);
				}
				beans.add(t);
				if (rowErrors.length() > 0) {
					errors.append(rowErrors).append("\r\n");
					rowErrors.setLength(0);
				}
			}
		} catch (Exception e) {
			if (e instanceof POIException) {
				errors.append(new StringBuffer(e.getMessage()).append("\t"));
				LOG.error(e.getMessage());
			} else {
				e.printStackTrace();
			}

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

	private static Object getCellValue(Row r, int cellNum) {
		// 缺失列处理政策
		Cell cell = r.getCell(cellNum, Row.CREATE_NULL_AS_BLANK);
		Object obj = null;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			obj = cell.getRichStringCellValue().getString();
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				obj = cell.getDateCellValue();
			} else {
				// 处理POI读取数字自动加.
				NumberFormat nf = NumberFormat.getInstance();
				String result = nf.format(cell.getNumericCellValue());
				if (result.indexOf(",") >= 0) {
					result = result.replace(",", "");
				}
				obj = result;
			}
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			obj = cell.getBooleanCellValue();
			break;
		case Cell.CELL_TYPE_FORMULA:
			obj = cell.getCellFormula();
			break;
		}
		return obj;
	}

	/**
	 * 获取Excel2003图片
	 * 
	 * @param sheetNum 当前sheet下标
	 * @param sheet    当前sheet对象
	 * @param workbook 工作簿对象
	 * @return Map key:图片单元格索引（0-sheet下标,1-列号,1-行号）String，value:图片流PictureData
	 * @throws IOException
	 */
	public static Map<String, PictureData> getSheetPictures(int sheetNum, Sheet sheet, Workbook workbook) {
		try {
			HSSFSheet sheetHSSF = (HSSFSheet) sheet;
			HSSFWorkbook workbookHSSF = (HSSFWorkbook) workbook;
			return getSheetPictrues03(sheetNum, sheetHSSF, workbookHSSF);
		} catch (Exception e) {
			XSSFSheet sheetXSSF = (XSSFSheet) sheet;
			XSSFWorkbook workbookXSSF = (XSSFWorkbook) workbook;
			return getSheetPictrues07(sheetNum, sheetXSSF, workbookXSSF);
		}
	}

	/**
	 * 获取Excel2003图片
	 * 
	 * @param sheetNum 当前sheet编号
	 * @param sheet    当前sheet对象
	 * @param workbook 工作簿对象
	 * @return Map key:图片单元格索引（0-sheet下标,1-列号,1-行号）String，value:图片流PictureData
	 * @throws IOException
	 */
	private static Map<String, PictureData> getSheetPictrues03(int sheetNum, HSSFSheet sheet, HSSFWorkbook workbook) {
		Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
		List<HSSFPictureData> pictures = workbook.getAllPictures();
		if (pictures.size() != 0) {
			for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
				HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
				if (shape instanceof HSSFPicture) {
					HSSFPicture pic = (HSSFPicture) shape;
					int pictureIndex = pic.getPictureIndex() - 1;
					HSSFPictureData picData = pictures.get(pictureIndex);
					String picIndex = String.valueOf(sheetNum) + "," + String.valueOf(anchor.getRow1()) + "," + String.valueOf(anchor.getCol1());
					sheetIndexPicMap.put(picIndex, picData);
				}
			}
			return sheetIndexPicMap;
		} else {
			return null;
		}
	}

	/**
	 * 获取Excel2007图片
	 * 
	 * @param sheetNum 当前sheet编号
	 * @param sheet    当前sheet对象
	 * @param workbook 工作簿对象
	 * @return Map key:图片单元格索引（0,1,1）String，value:图片流PictureData
	 */
	private static Map<String, PictureData> getSheetPictrues07(int sheetNum, XSSFSheet sheet, XSSFWorkbook workbook) {
		Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
		for (POIXMLDocumentPart dr : sheet.getRelations()) {
			if (dr instanceof XSSFDrawing) {
				XSSFDrawing drawing = (XSSFDrawing) dr;
				List<XSSFShape> shapes = drawing.getShapes();
				for (XSSFShape shape : shapes) {
					if (shape instanceof XSSFPicture) {
						XSSFPicture pic = (XSSFPicture) shape;
						XSSFClientAnchor anchor = pic.getPreferredSize();
						CTMarker ctMarker = anchor.getFrom();
						String picIndex = String.valueOf(sheetNum) + "," + ctMarker.getRow() + "," + ctMarker.getCol();
						sheetIndexPicMap.put(picIndex, pic.getPictureData());
					}
				}
			}
		}
		return sheetIndexPicMap;
	}

	/**
	 * 
	 * excel导出头类文件 ,如果表头合并，只需设置hearderRules，无需设置title,hearder,
	 *
	 * @author lt
	 * @version 2017年11月21日上午9:45:01
	 */
	public static class ExportRules {

		/**
		 * 是否带序号
		 */

		private boolean autoNum;

		/**
		 * 导出字段设置，
		 * 
		 */

		private Object[][] fields;

		/**
		 * 表头名
		 */
		private String titile;

		/**
		 * 标题列
		 */
		private String[] header;

		/**
		 * excel头：合并规则及值，rules.put("1,1,A,G", "其它应扣"); 对应excel位置
		 */
		private HashMap<String, String> headerRules;

		/**
		 * excel尾 ： 合并规则及值，rules.put("1,1,A,G", 值); 对应excel位置
		 */
		private HashMap<String, String> footerRules;

		// --------------------无关设置字段-------------------------

		/**
		 * 最大单元格列数
		 */
		private int maxColumns = 0;

		/**
		 * 表头最大行数
		 */
		private int maxRows = 0;

		/**
		 * 是否合并表头
		 */
		private boolean ifMerge;

		/**
		 * 是否有页脚
		 */
		private boolean ifFooter;

		/**
		 * 常规一行表头构造,不带尾部
		 * 
		 * @param autoNum     是否自动序号
		 * @param fields      二维数组，1数据字段，2列宽，3居左居右（非必），如Object[][] fileds = { { "projectName", POIConstant.AUTO,POIConstant.RIGHT }
		 * @param titile      大标题
		 * @param header      表头标题
		 * @param footerRules 数据尾行合计
		 */
		public ExportRules(boolean autoNum, Object[][] fields, String titile, String[] header, HashMap<String, String> footerRules) {
			super();
			this.autoNum = autoNum;
			this.fields = fields;
			if (titile != null) {
				setTitile(titile);
			}
			setHeader(header);
			if (footerRules != null) {
				setFooterRules(footerRules);
			}
		}

		/**
		 * 复杂表头构造
		 * 
		 * @param autoNum     是否自动序号
		 * @param fields      二维数组，1数据字段，2列宽，3居左居右（非必），如Object[][] fileds = { { "projectName", POIConstant.AUTO,POIConstant.RIGHT }
		 * @param headerRules 表头设计
		 * @param footerRules 尾部合计栏设计
		 */
		public ExportRules(boolean autoNum, Object[][] fields, HashMap<String, String> headerRules, HashMap<String, String> footerRules) {
			super();
			this.autoNum = autoNum;
			this.fields = fields;
			setHeaderRules(headerRules);
			if (footerRules != null) {
				setFooterRules(footerRules);
			}
		}

		public boolean getIfFooter() {
			return ifFooter;
		}

		public boolean getIfMerge() {
			return ifMerge;
		}

		public String getTitile() {
			return titile;
		}

		private void setTitile(String titile) {
			this.ifMerge = false;
			this.titile = titile;
			this.maxRows = this.maxRows + 1;
		}

		public String[] getHearder() {
			return header;
		}

		private void setHeader(String[] header) {
			this.ifMerge = false;
			this.header = header;
			this.maxRows = this.maxRows + 1;
			this.maxColumns = header.length - 1;
		}

		public int getMaxColumns() {
			return maxColumns;
		}

		public int getMaxRows() {
			return maxRows;
		}

		public HashMap<String, String> getHeaderRules() {
			return headerRules;
		}

		private void setHeaderRules(HashMap<String, String> headerRules) {
			this.headerRules = headerRules;
			// 解析rules，获取最大行和最大列
			Iterator<Entry<String, String>> entries = headerRules.entrySet().iterator();
			int row = 0;
			int col = 0;
			while (entries.hasNext()) {
				Entry<String, String> entry = entries.next();
				String key = entry.getKey();
				Object[] range = coverRange(key);
				int a = (int) range[1];
				int b = POIConstant.cellRefNums.get(range[3]) + 1;
				row = a > row ? a : row;
				col = b > col ? b : col;
			}
			this.maxRows = row;
			this.maxColumns = col;
			this.ifMerge = true;
		}

		/**
		 * 合并单元格转换
		 * 
		 * @param obj
		 * @return
		 */
		private static Object[] coverRange(Object obj) {
			String[] range = ((String) obj).split(",");
			Object[] rangeInt = new Object[4];
			for (int i = 0; i < range.length; i++) {
				if (i < 2) {
					rangeInt[i] = Integer.parseInt(range[i]);
				} else {
					rangeInt[i] = range[i];
				}

			}
			return rangeInt;
		}

		public HashMap<String, String> getFooterRules() {
			return footerRules;
		}

		private void setFooterRules(HashMap<String, String> footerRules) {
			this.ifFooter = true;
			this.footerRules = footerRules;
		}

		public boolean getAutoNum() {
			return autoNum;
		}

		public Object[][] getFields() {
			return fields;
		}

	}

	/**
	 * 将流转换为byte数组，作为图片数据导入
	 * 
	 * @param fis
	 * @return
	 */
	public static byte[] ImageParseBytes(InputStream fis) {
		byte[] buffer = null;
		ByteArrayOutputStream bos = null;
		try {
			bos = new ByteArrayOutputStream(1024);
			byte[] b = new byte[1024];
			int n;
			while ((n = fis.read(b)) != -1) {
				bos.write(b, 0, n);
			}
			buffer = bos.toByteArray();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				fis.close();
				bos.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return buffer;
	}

	public static byte[] ImageParseBytes(File file) {
		FileInputStream fileInputStream = null;
		try {
			fileInputStream = new FileInputStream(file);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		return ImageParseBytes(fileInputStream);
	}
}
