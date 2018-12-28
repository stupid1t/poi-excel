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
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import excel.callBack.ExportSheetCallback;
import excel.callBack.ParseSheetCallback;
import excel.verify.AbstractCellVerify;
import excel.verify.AbstractVerifyBuidler;
import excel.verify.ImgVerify;

/**
 * excel导入导出工具，也可以导出模板
 * 
 * @author 625
 *
 */
public class ExcelUtils {

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
		titleFont.setBold(true);
		style = wb.createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);// 左右居中
		style.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
		style.setFont(titleFont);
		styles.put("title", style);

		Font monthFont = wb.createFont();
		monthFont.setFontName("Arial");
		monthFont.setFontHeightInPoints((short) 10);
		monthFont.setColor(IndexedColors.WHITE.getIndex());
		style = wb.createCellStyle();
		style.setBorderRight(BorderStyle.THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(BorderStyle.THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(monthFont);
		style.setWrapText(true);
		styles.put("header", style);

		style = wb.createCellStyle();
		Font cellFont = wb.createFont();
		cellFont.setFontName("Arial");
		cellFont.setFontHeightInPoints((short) 10);
		style.setFont(cellFont);
		style.setWrapText(false);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setBorderRight(BorderStyle.THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(BorderStyle.THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		styles.put("cell", style);

		return styles;
	}

	/**
	 * 导出
	 * 
	 * @param             <T>
	 * 
	 * @param data        数据源
	 * @param exportRules 导出规则
	 * @return
	 */
	public static <T> Workbook createWorkbook(List<T> data, ExportRules exportRules) {
		Workbook work = createWorkbook(data, exportRules, null);
		return work;
	}

	/**
	 * 
	 * 导出方法
	 * 
	 * @param list         数据源
	 * @param hearderRules （如带序号，在规则里设置序头） 表头规则
	 * @param autoNum      带序号
	 * @param column       导出列的定义
	 * @return
	 * @throws Exception
	 */
	public static <T> Workbook createWorkbook(List<T> list, ExportRules hearderRules, ExportSheetCallback<T> callBack) {
		boolean autoNum = hearderRules.getAutoNum();
		Column[] fields = hearderRules.getColumn();
		boolean is07 = hearderRules.isXlsx();
		Workbook wb = null;
		if (is07) {
			wb = new XSSFWorkbook();// 2007
		} else {
			wb = new HSSFWorkbook();// 2003
		}
		Map<String, CellStyle> styles = ExcelUtils.initStyles(wb);
		CellStyle titleStyle = styles.get("title");
		CellStyle headerStyle = styles.get("header");
		CellStyle cellStyle = styles.get("cell");
		Sheet sheet = wb.createSheet();
		ExcelUtils.printSetup(sheet);
		// -----------------------表头设置------------------------
		int maxColumns = hearderRules.getMaxColumns();
		int maxRows = hearderRules.getMaxRows();

		// 创建表头
		for (int i = 0; i < maxRows; i++) {
			Row row = sheet.createRow(i);
			for (int j = 0; j < maxColumns; j++) {
				row.createCell(j);
			}
		}

		if (hearderRules.getIfMerge()) {// 合并模式
			// 冻结表头
			sheet.createFreezePane(0, maxRows, 0, maxRows);
			// header
			Map<String, String> rules = hearderRules.getHeaderRules();
			Iterator<Entry<String, String>> entries = rules.entrySet().iterator();
			while (entries.hasNext()) {
				Entry<String, String> entry = entries.next();
				String key = entry.getKey();
				String value = entry.getValue();
				Object[] range = coverRange(key);
				// 合并表头
				int firstRow = (int) range[0] - 1;
				int lastRow = (int) range[1] - 1;
				int firstCol = POIConstant.cellRefNums.get(range[2]);
				int lastCol = POIConstant.cellRefNums.get(range[3]);
				if ((lastRow - firstRow) != 0 || (lastCol - firstCol) != 0) {
					CellRangeAddress cra = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
					sheet.addMergedRegion(cra);
					RegionUtil.setBorderTop(BorderStyle.THIN, cra, sheet);
					RegionUtil.setBorderBottom(BorderStyle.THIN, cra, sheet);
					RegionUtil.setBorderLeft(BorderStyle.THIN, cra, sheet);
					RegionUtil.setBorderRight(BorderStyle.THIN, cra, sheet);
				}

				if ((maxColumns - 1) == lastCol - firstCol) {// 占满全格，则为表头
					CellUtil.createCell(sheet.getRow(firstRow), firstCol, value, titleStyle);
				} else {
					CellUtil.createCell(sheet.getRow(firstRow), firstCol, value, headerStyle);
				}
			}
		} else {// 非合并
			if (hearderRules.getTitile() == null) {
				// 冻结表头
				sheet.createFreezePane(0, 1, 0, 1);
				String[] hearder = hearderRules.getHearder();
				for (int i = 0; i < hearder.length; i++) {
					CellUtil.createCell(sheet.getRow(0), i, hearder[i], headerStyle);
				}
			} else {
				// 冻结表头
				sheet.createFreezePane(0, 2, 0, 2);
				CellRangeAddress cra = new CellRangeAddress(0, 0, 0, maxColumns);
				sheet.addMergedRegion(cra);
				CellUtil.createCell(sheet.getRow(0), 0, hearderRules.getTitile(), titleStyle);
				String[] hearder = hearderRules.getHearder();
				for (int i = 0; i < hearder.length; i++) {
					CellUtil.createCell(sheet.getRow(1), i, hearder[i], headerStyle);
				}
			}

		}
		// --------------------set width--------------------
		for (int i = 0, j = 0; i < fields.length; i++, j++) {
			// 0.每一列默认单元格样式设置
			// 1.width设置
			if (autoNum && j == 0) {
				j++;
				sheet.setColumnWidth(0, 2000);
			}
			Column column = fields[i];
			// 1.1是否自动列宽
			int width = column.getWidth();
			if (width != 0) {
				sheet.setColumnWidth(j, width);
			} else {
				// 1.2根据maxRows，获取表头的值设置宽度
				Row row = sheet.getRow(maxRows - 1);
				String headerValue = row.getCell(j).getStringCellValue();
				if (StringUtils.isBlank(headerValue)) {
					row = sheet.getRow(maxRows - 2);
					headerValue = row.getCell(j).getStringCellValue();
				}
				sheet.setColumnWidth(j, headerValue.getBytes().length * 256);
			}
			// 2.downDown设置
			String[] dorpDown = column.getDorpDown();
			if (dorpDown != null && dorpDown.length > 0) {
				sheet.addValidationData(createDropDownValidation(sheet, dorpDown, j, maxRows));
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
					throw new IllegalArgumentException("时间校验表达式不正确,请填写如2015-08-09~2016-09-10的值!");
				}
				try {
					sheet.addValidationData(createDateValidation(sheet, split1[0], split1[1], info, j, maxRows));
				} catch (Exception e) {
					e.printStackTrace();
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
				sheet.addValidationData(createNumValidation(sheet, split1[0], split1[1], info, j, maxRows));
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
				sheet.addValidationData(createFloatValidation(sheet, split1[0], split1[1], info, j, maxRows));
			}

			// 5.自定义校验
			String custom = column.getVerifyCustom();
			if (custom != null) {
				String[] split = custom.split("@");
				String info = null;
				if (split.length == 2) {
					info = split[1];
				}
				sheet.addValidationData(createCustomValidation(sheet, split[0], info, j, maxRows));
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
				sheet.addValidationData(createTextLengthValidation(sheet, split2[0], split2[1], info, j, maxRows));
			}
		}

		// ------------------body row-----------------
		// 画图器
		@SuppressWarnings("unchecked")
		Drawing<Picture> createDrawingPatriarch = (Drawing<Picture>) sheet.createDrawingPatriarch();
		// 存储类的字段信息
		Map<Class<? extends Object>, Map<String, Field>> clsInfo = new HashMap<>();
		// 存储单元格样式信息，此方式与因为POI的一个BUG
		Map<Object, CellStyle> subCellStyle = new HashMap<>();
		for (int i = 0; i < list.size(); i++) {
			Row row = sheet.createRow(i + maxRows);
			T t = list.get(i);
			for (int j = 0, n = 0; n < fields.length; j++, n++) {
				Cell cell = row.createCell(j);
				cell.setCellStyle(cellStyle);
				// 1.序号设置
				if (autoNum && j == 0) {
					cell.setCellValue(i + 1);
					n--;
					continue;
				}
				// 2.读取Map/Object对应字段值
				if (clsInfo.get(t.getClass()) == null) {
					clsInfo.put(t.getClass(), getAllFields(t.getClass()));
				}
				Object value = readField(clsInfo, t, fields[n].getField());

				// 3.填充列值
				Column customStyle = null;
				if (callBack != null) {
					customStyle = Column.style();
					value = callBack.callback(fields[n].getField(), value, t, customStyle);
				}
				// 4.设置单元格值
				setCellValue(createDrawingPatriarch, fields[n], customStyle, value, cell, subCellStyle);
			}
		}
		// ------------------------footer row-----------------------------
		if (hearderRules.getIfFooter()) {
			Map<String, String> footerRules = hearderRules.getFooterRules();
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
				int firstRow = (int) range[0] + currRownum - 1;
				int lastRow = (int) range[1] + currRownum - 1;
				int firstCol = POIConstant.cellRefNums.get(range[2]);
				int lastCol = POIConstant.cellRefNums.get(range[3]);
				if ((lastRow - firstRow) != 0 || (lastCol - firstCol) != 0) {
					CellRangeAddress cra = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
					sheet.addMergedRegion(cra);
					RegionUtil.setBorderBottom(BorderStyle.THIN, cra, sheet);
					RegionUtil.setBorderTop(BorderStyle.THIN, cra, sheet);
					RegionUtil.setBorderLeft(BorderStyle.THIN, cra, sheet);
					RegionUtil.setBorderRight(BorderStyle.THIN, cra, sheet);
				}
				String cellValue = "";
				cellValue = value;
				CellUtil.createCell(sheet.getRow(firstRow), firstCol, cellValue, cellStyle);
			}

		}
		return wb;
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
		// 获取excel中所有图片
		List<String> imgField = new ArrayList<>();
		Map<String, PictureData> pictures = null;
		Map<String, AbstractCellVerify> verifys = verifyBuilder.getVerifys();
		Set<String> keySet = verifys.keySet();
		int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
		for (String key : keySet) {
			AbstractCellVerify abstractCellVerify = verifys.get(key);
			if (abstractCellVerify instanceof ImgVerify) {
				imgField.add(key);
				if (pictures == null) {
					pictures = getSheetPictures(sheetIndex, sheet);
				}
			}
		}
		StringBuffer errors = new StringBuffer();
		StringBuffer rowErrors = new StringBuffer();
		try {
			int rowStart = sheet.getFirstRowNum() + dataStartRow;
			// warn获取真实的数据行尾数
			int rowEnd = getLastRealLastRow(sheet.getRow(sheet.getLastRowNum())) - dataEndRowCount;
			for (int rowNum = rowStart; rowNum <= rowEnd; rowNum++) {
				Row r = sheet.getRow(rowNum);
				// 创建对象
				T t = clss.newInstance();
				int fieldNum = 0;
				for (int cellNum : POIConstant.convertToCellNum(verifyBuilder.cellRefs)) {
					// 列坐标
					CellReference cellRef = new CellReference(rowNum, cellNum);
					String filedName = verifyBuilder.filedNames[fieldNum];
					try {
						Object cellValue = null;
						if (imgField.size() > 0 && imgField.contains(filedName)) {
							String pictrueIndex = sheetIndex + "," + rowNum + "," + cellNum;
							PictureData remove = pictures.remove(pictrueIndex);
							cellValue = remove == null ? null : remove.getData();
						} else {
							cellValue = getCellValue(r, cellNum);
						}
						// 校验和格式化列值
						cellValue = verifyBuilder.verify(filedName, cellValue);
						// 填充列值
						FieldUtils.writeDeclaredField(t, filedName, cellValue, true);
					} catch (POIException e) {
						rowErrors.append(cellRef.formatAsString()).append(":").append(e.getMessage()).append("\t");
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

	/**
	 * 获取真实的数据行
	 * 
	 * @param r
	 * @return
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
				if (cell.getCellType() == CellType.BLANK) {
					blankCell++;
				}
			}
			if (blankCell >= lastCellNum) {
				int rowNum = row.getRowNum();
				Row newRow = sheet.getRow(--rowNum);
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
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings("rawtypes")
	private static Object readField(Map<Class<? extends Object>, Map<String, Field>> clsInfo, Object t, String fields) {
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
	 * @param wb                     工作簿
	 * @param value                  列值
	 * @param pattern                格式化值
	 * @param cell                   单元格
	 * @param createDrawingPatriarch 画图器
	 * @param subCellStyle
	 */
	private static void setCellValue(Drawing<Picture> createDrawingPatriarch, Column sourceColumn, Column customColumn, Object value, Cell cell, Map<Object, CellStyle> subCellStyle) {
		Workbook workbook = cell.getSheet().getWorkbook();
		// 0.判断是否需要用用户的样式
		boolean customer = false;
		if (customColumn != null) {
			customer = (customColumn.set == 1);
		}
		// 1.水平定位
		HorizontalAlignment align = customer ? customColumn.getAlign() : sourceColumn.getAlign();
		if (align != null) {
			// 表示需要用户自定义的定位
			CellStyle style = subCellStyle.get(customer + "-align-" + align);
			if (style == null) {
				CellStyle sourceStyle = cell.getCellStyle();
				style = workbook.createCellStyle();
				style.cloneStyleFrom(sourceStyle);
				style.setAlignment(align);
				subCellStyle.put(customer + "-align-" + align, style);
			}
			cell.setCellStyle(style);
		}
		// 2.垂直定位
		VerticalAlignment valign = customer ? customColumn.getValign() : sourceColumn.getValign();
		if (valign != null) {
			// 表示需要用户自定义的定位
			CellStyle style = subCellStyle.get(customer + "-valign-" + valign);
			if (style == null) {
				CellStyle sourceStyle = cell.getCellStyle();
				style = workbook.createCellStyle();
				style.cloneStyleFrom(sourceStyle);
				style.setVerticalAlignment(valign);
				subCellStyle.put(customer + "-valign-" + valign, style);
			}
			cell.setCellStyle(style);
		}
		// 3.字体颜色
		IndexedColors color = customer ? customColumn.getColor() : sourceColumn.getColor();
		if (color != null) {
			// 表示需要用户自定义的定位
			CellStyle style = subCellStyle.get(customer + "-color-" + color);
			if (style == null) {
				CellStyle sourceStyle = cell.getCellStyle();
				style = workbook.createCellStyle();
				style.cloneStyleFrom(sourceStyle);
				Font font = workbook.createFont();
				font.setFontName("Arial");
				font.setFontHeightInPoints((short) 10);
				font.setColor(color.getIndex());
				style.setFont(font);
				subCellStyle.put(customer + "-color-" + color, style);
			}
			cell.setCellStyle(style);
		}
		// 4.背景色
		IndexedColors backColor = customer ? customColumn.getBackColor() : sourceColumn.getBackColor();
		if (backColor != null) {
			// 表示需要用户自定义的定位
			CellStyle style = subCellStyle.get(customer + "-backColor-" + backColor);
			if (style == null) {
				CellStyle sourceStyle = cell.getCellStyle();
				style = workbook.createCellStyle();
				style.cloneStyleFrom(sourceStyle);
				style.setFillForegroundColor(backColor.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				subCellStyle.put(customer + "-backColor-" + backColor, style);
			}
			cell.setCellStyle(style);
		}

		// 4.高度
		int height = customer ? customColumn.getHeight() : sourceColumn.getHeight();
		if (height != 0) {
			// 表示需要用户自定义高度
			cell.getRow().setHeight((short) height);
		}

		// 判断值的类型后进行强制类型转换.再设置单元格格式
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
			// 1.格式化为年月日时分
			String pattern = POIConstant.FMTDATETIME;
			// 2.判断时分秒是否为0，如果是格式化为年月日
			Date date = (Date) value;
			Calendar cal = Calendar.getInstance();
			cal.setTime(date);
			int hour = cal.get(Calendar.HOUR);
			int minute = cal.get(Calendar.MINUTE);
			int second = cal.get(Calendar.SECOND);
			if ((hour - minute - second) == 0) {
				pattern = POIConstant.FMTDATE;
			}
			CellStyle style = subCellStyle.get(pattern);
			if (style == null) {
				CellStyle sourceStyle = cell.getCellStyle();
				style = workbook.createCellStyle();
				style.cloneStyleFrom(sourceStyle);
				CreationHelper createHelper = workbook.getCreationHelper();
				style.setDataFormat(createHelper.createDataFormat().getFormat(pattern));
				subCellStyle.put(pattern, style);
			}
			cell.setCellStyle(style);
			cell.setCellValue(date);
		} else if (value instanceof byte[]) {
			byte[] data = (byte[]) value;
			// 5.1anchor主要用于设置图片的属性
			short x = (short) cell.getColumnIndex();
			int y = cell.getRowIndex();
			// 5.2插入图片
			ClientAnchor anchor = null;
			if (workbook instanceof XSSFWorkbook) {
				anchor = new XSSFClientAnchor(10, 10, 10, 10, x, y, x + 1, y + 1);
			} else {
				anchor = new HSSFClientAnchor(10, 10, 10, 10, (short) x, y, (short) (x + 1), y + 1);
			}
			int add1 = workbook.addPicture(data, XSSFWorkbook.PICTURE_TYPE_PNG);
			createDrawingPatriarch.createPicture(anchor, add1);
			cell.setCellValue("");
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
	 * 获取单元格的值
	 * 
	 * @param r
	 * @param cellNum
	 * @return
	 */
	private static Object getCellValue(Row r, int cellNum) {
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
				// 处理POI读取数字自动加.
				NumberFormat nf = NumberFormat.getInstance();
				String result = nf.format(cell.getNumericCellValue());
				if (result.indexOf(",") >= 0) {
					result = result.replace(",", "");
				}
				obj = result;
			}
			break;
		case BOOLEAN:
			obj = cell.getBooleanCellValue();
			break;
		case FORMULA:
			obj = cell.getCellFormula();
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
	 * @throws IOException
	 */
	private static Map<String, PictureData> getSheetPictures(int sheetNum, Sheet sheet) {
		try {
			HSSFSheet sheetHSSF = (HSSFSheet) sheet;
			return getSheetPictrues03(sheetNum, sheetHSSF);
		} catch (Exception e) {
			XSSFSheet sheetXSSF = (XSSFSheet) sheet;
			return getSheetPictrues07(sheetNum, sheetXSSF);
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
	private static Map<String, PictureData> getSheetPictrues03(int sheetNum, HSSFSheet sheet) {
		Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
		List<HSSFPictureData> pictures = sheet.getWorkbook().getAllPictures();
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
	private static Map<String, PictureData> getSheetPictrues07(int sheetNum, XSSFSheet sheet) {
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
	 * excel添加下拉数据校验
	 * 
	 * 
	 * @param sheet      哪个 sheet 页添加校验
	 * @param dataSource 数据源数组
	 * @param col        第几列校验（0开始）
	 * @param maxRow     表头占用几行
	 * @return
	 */
	private static DataValidation createDropDownValidation(Sheet sheet, String[] dataSource, int col, int maxRow) {
		CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(maxRow, 65535, col, col);
		DataValidationHelper helper = sheet.getDataValidationHelper();
		DataValidationConstraint constraint = null;
		if (dataSource.length < 11) {
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
			if (rowNum == 0) {
				// 第一次创建下拉框数据
				for (int i = 0; i < dataLength; i++, rowNum++) {
					hidden.createRow(i).createCell(0).setCellValue(dataSource[i]);
				}
			} else {
				// 之前已经创建过
				int createNum = dataLength - ++rowNum;
				short lastCellNum = (short) (hidden.getRow(0).getLastCellNum());
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
	 * 
	 * @param sheet      哪个 sheet 页添加校验
	 * @param dataSource 数据源数组
	 * @param col        第几列校验（0开始）
	 * @param maxRow     表头占用几行
	 * @return
	 */
	private static DataValidation createDateValidation(Sheet sheet, String start, String end, String info, int col, int maxRow) throws Exception {
		String pattern = POIConstant.FMTDATETIME;
		// 0.格式判断
		if (start.length() != 16) {
			pattern = POIConstant.FMTDATE;
		}
		// 1.设置验证
		CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(maxRow, 65535, col, col);
		DataValidationHelper helper = sheet.getDataValidationHelper();
		Calendar cal = Calendar.getInstance();
		Date startDate = DateUtils.parseDate(start, pattern);
		Date endDate = DateUtils.parseDate(end, pattern);
		cal.setTime(startDate);
		String formulaStart = "=DATE(" + cal.get(Calendar.YEAR) + "," + (cal.get(Calendar.MONTH) + 1) + "," + cal.get(Calendar.DATE) + ")";
		cal.setTime(endDate);
		String formulaEnd = "=DATE(" + cal.get(Calendar.YEAR) + "," + (cal.get(Calendar.MONTH) + 1) + "," + cal.get(Calendar.DATE) + ")";
		DataValidationConstraint constraint = helper.createDateConstraint(OperatorType.BETWEEN, formulaStart, formulaEnd, pattern);
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
		// 2.设置单元格格式
		Workbook workbook = sheet.getWorkbook();
		CellStyle style = workbook.createCellStyle();
		CreationHelper createHelper = workbook.getCreationHelper();
		style.setDataFormat(createHelper.createDataFormat().getFormat(pattern));
		sheet.setDefaultColumnStyle(col, style);
		return dataValidation;
	}

	/**
	 * excel添加数字校验
	 * 
	 * 
	 * @param sheet      哪个 sheet 页添加校验
	 * @param dataSource 数据源数组
	 * @param col        第几列校验（0开始）
	 * @param maxRow     表头占用几行
	 * @return
	 */
	private static DataValidation createNumValidation(Sheet sheet, String minNum, String maxNum, String info, int col, int maxRow) {
		// 1.设置验证
		CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(maxRow, 65535, col, col);
		DataValidationHelper helper = sheet.getDataValidationHelper();
		DataValidationConstraint constraint = helper.createIntegerConstraint(OperatorType.BETWEEN, minNum, maxNum);
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
	 * 
	 * @param sheet  哪个 sheet 页添加校验
	 * @param minNum 最小值
	 * @param maxNum 最大值
	 * @param col    第几列校验（0开始）
	 * @param maxRow 表头占用几行
	 * @return
	 */
	private static DataValidation createFloatValidation(Sheet sheet, String minNum, String maxNum, String info, int col, int maxRow) {
		// 1.设置验证
		CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(maxRow, 65535, col, col);
		DataValidationHelper helper = sheet.getDataValidationHelper();
		DataValidationConstraint constraint = helper.createDecimalConstraint(OperatorType.BETWEEN, minNum, maxNum);
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
	 * excel添加文本字符长度校验
	 * 
	 * 
	 * @param sheet  哪个 sheet 页添加校验
	 * @param minNum 最小值
	 * @param maxNum 最大值
	 * @param info   自定义提示
	 * @param col    第几列校验（0开始）
	 * @param maxRow 表头占用几行
	 * @return
	 */
	private static DataValidation createTextLengthValidation(Sheet sheet, String minNum, String maxNum, String info, int col, int maxRow) {
		// 1.设置验证
		CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(maxRow, 65535, col, col);
		DataValidationHelper helper = sheet.getDataValidationHelper();
		DataValidationConstraint constraint = helper.createTextLengthConstraint(OperatorType.BETWEEN, minNum, maxNum);
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
	 * excel添加自定义校验
	 * 
	 * 
	 * @param sheet   哪个 sheet 页添加校验
	 * @param formula 表达式
	 * @param col     第几列校验（0开始）
	 * @param maxRow  表头占用几行
	 * @return
	 */
	private static DataValidation createCustomValidation(Sheet sheet, String formula, String info, int col, int maxRow) {
		String msg = "请输入正确的值！";
		// 0.修正xls表达式不正确定位的问题,只修正了开始，如F3:F2000,修正了F3变为A0,F2000变为A2000
		Workbook workbook = sheet.getWorkbook();
		if (workbook instanceof HSSFWorkbook) {
			// 替换字母为A，下标从0开始
			int start = formula.indexOf("(") + 1;
			int end = formula.indexOf(")");
			if (start != 1 && end != 0) {
				String prev = formula.substring(0, start);
				String sufix = formula.substring(end, formula.length());
				String substring = formula.substring(start, end);
				char[] charArray = substring.toCharArray();
				int over = 0;
				for (int i = 0; i < charArray.length; i++) {
					char c = charArray[i];
					if (c == ':') {
						over++;
						continue;
					}
					if (!Character.isDigit(c)) {
						charArray[i] = 'A';
					} else {
						if (over == 0) {
							charArray[i] = String.valueOf(maxRow - 1).charAt(0);
						}
					}
				}
				formula = prev + String.valueOf(charArray) + sufix;
			}

		}
		// 1.设置验证
		CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(maxRow, 65535, col, col);
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
		} catch (Exception e) {
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

	/**
	 * 将文件转换为byte数组，作为图片数据导入
	 * 
	 * @param fis
	 * @return
	 */
	public static byte[] ImageParseBytes(File file) {
		FileInputStream fileInputStream = null;
		try {
			fileInputStream = new FileInputStream(file);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		return ImageParseBytes(fileInputStream);
	}

	/**
	 * 
	 * excel导出规则
	 *
	 * @author lt
	 * @version 2017年11月21日上午9:45:01
	 */
	public static class ExportRules {

		/**
		 * 是否带序号
		 */

		private boolean xlsx = false;

		/**
		 * 是否带序号
		 */

		private boolean autoNum;

		/**
		 * 列数据规则定义
		 */
		private Column[] column;

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
		private Map<String, String> headerRules;

		/**
		 * excel尾 ： 合并规则及值，rules.put("1,1,A,G", 值); 对应excel位置
		 */
		private Map<String, String> footerRules;

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
		 * @param fields      列数据规则定义
		 * @param titile      大标题
		 * @param header      表头标题
		 * @param footerRules 数据尾行合计
		 */
		public ExportRules(boolean autoNum, Column[] column, String titile, String[] header, HashMap<String, String> footerRules) {
			super();
			this.autoNum = autoNum;
			this.column = column;
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
		 * @param fields      列数据规则定义
		 * @param headerRules 表头设计
		 * @param footerRules 尾部合计栏设计
		 */
		public ExportRules(boolean autoNum, Column[] column, Map<String, String> headerRules, Map<String, String> footerRules) {
			super();
			this.autoNum = autoNum;
			this.column = column;
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

		public Map<String, String> getHeaderRules() {
			return headerRules;
		}

		private void setHeaderRules(Map<String, String> headerRules) {
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

		public Map<String, String> getFooterRules() {
			return footerRules;
		}

		private void setFooterRules(Map<String, String> footerRules) {
			this.ifFooter = true;
			this.footerRules = footerRules;
		}

		public boolean getAutoNum() {
			return autoNum;
		}

		public Column[] getColumn() {
			return column;
		}

		public boolean isXlsx() {
			return xlsx;
		}

		public ExportRules setXlsx(boolean xlsx) {
			this.xlsx = xlsx;
			return this;
		}

	}

	/**
	 * 列的定义
	 * 
	 * @author 625
	 *
	 */
	public static class Column {

		/**
		 * 字段名称
		 */
		private String field;

		/**
		 * 宽度，不设置默认自动
		 */
		private int width;

		/**
		 * 高度，设置是行的高度
		 */
		private int height;

		/**
		 * 水平定位，默认居中
		 */
		private HorizontalAlignment align;

		/**
		 * 垂直定位，默认居下
		 */
		private VerticalAlignment valign;

		/**
		 * 字体颜色，默认黑色
		 */
		private IndexedColors color;

		/**
		 * 背景色，默认无
		 */
		private IndexedColors backColor;

		/**
		 * 下拉列表数据
		 */
		private String[] dorpDown;

		/**
		 * 日期校验,请填写例如2018-08-09~2019-08-09
		 */
		private String verifyDate;

		/**
		 * 整数数字校验,请填写例如10~30
		 */
		private String verifyIntNum;

		/**
		 * 浮点数字校验,请填写例如10.0~30.0
		 */
		private String verifyFloatNum;

		/**
		 * 文本长度校验
		 */
		private String verifyText;

		/**
		 * 自定义表达式校验
		 */
		private String verifyCustom;

		/**
		 * 定义规则个数
		 */
		private int verifyCount;

		/**
		 * 是否为回调样式模式
		 */
		private int style;

		/**
		 * 判断用户是否重置样式
		 */
		private int set;

		private Column(String field) {
			this.field = field;
		}

		private Column() {

		}

		/**
		 * 字段名称
		 * 
		 * @param field 字段名称
		 * @return
		 */
		public static Column style() {
			Column column = new Column();
			column.style = 1;
			return column;
		}

		/**
		 * 字段名称
		 * 
		 * @param field 字段名称
		 * @return
		 */
		public static Column field(String field) {
			return new Column(field);
		}

		String getField() {
			return field;
		}

		int getHeight() {
			return height;
		}

		/**
		 * 高度
		 * 
		 * @param height 不设置默认
		 * @return
		 */
		public Column height(int height) {
			if (style == 1) {
				set = 1;
			}
			this.height = POIConstant.width(height);
			return this;
		}

		HorizontalAlignment getAlign() {
			return align;
		}

		int getWidth() {
			return width;
		}

		/**
		 * 宽度
		 * 
		 * @param width 不设置默认自动
		 * @return
		 */
		public Column width(int width) {
			if (style == 1) {
				throw new UnsupportedOperationException("仅允许定义color/backColor/align/valign ！");
			}
			this.width = POIConstant.width(width);
			return this;
		}

		/**
		 * 水平定位
		 * 
		 * @param align ，CellStyle 取值
		 * @return
		 */
		public Column align(HorizontalAlignment align) {
			if (style == 1) {
				set = 1;
			}
			this.align = align;
			return this;
		}

		IndexedColors getColor() {
			return color;
		}

		/**
		 * 设置字体颜色
		 * 
		 * @param backColor HSSFColor,XSSFColor
		 * @return
		 */
		public Column color(IndexedColors color) {
			if (style == 1) {
				set = 1;
			}
			this.color = color;
			return this;
		}

		IndexedColors getBackColor() {
			return backColor;
		}

		/**
		 * 设置背景色
		 * 
		 * @param backColor
		 * @return
		 */
		public Column backColor(IndexedColors backColor) {
			if (style == 1) {
				set = 1;
			}
			this.backColor = backColor;
			return this;
		}

		VerticalAlignment getValign() {
			return valign;
		}

		/**
		 * 设置垂直定位
		 * 
		 * @param vAlign 默认居下
		 * @return
		 */
		public Column valign(VerticalAlignment valign) {
			if (style == 1) {
				set = 1;
			}
			this.valign = valign;
			return this;
		}

		String[] getDorpDown() {
			return dorpDown;
		}

		/**
		 * 下拉列表数据
		 * 
		 * @param dorpDown 下拉列表数据
		 * @return
		 */
		public Column dorpDown(String[] dorpDown) {
			if (style == 1) {
				throw new UnsupportedOperationException("仅允许定义color/backColor/align/valign ！");
			}
			if (++verifyCount > 1) {
				throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
			}
			this.dorpDown = dorpDown;
			return this;
		}

		String getVerifyDate() {
			return verifyDate;
		}

		/**
		 * 
		 * 日期数据校验
		 *
		 * @param verifyDate 表达式，请填写例如2018-08-09~2019-08-09 格式也可以 yyyy-MM-dd HH:mm:ss
		 * @param msgInfo    提示消息
		 * @return
		 */
		public Column verifyDate(String verifyDate, String... msgInfo) {
			if (style == 1) {
				throw new UnsupportedOperationException("仅允许定义color/backColor/align/valign ！");
			}
			if (++verifyCount > 1) {
				throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
			}
			if (msgInfo.length > 0) {
				this.verifyDate = verifyDate + "@" + msgInfo[0];
			} else {
				this.verifyDate = verifyDate;
			}

			return this;
		}

		String getVerifyIntNum() {
			return verifyIntNum;
		}

		/**
		 * 整数数字数据校验
		 * 
		 * @param verifyNum 表达式,请填写例如10~30
		 * @param msgInfo   提示消息
		 * @return
		 */
		public Column verifyIntNum(String verifyIntNum, String... msgInfo) {
			if (style == 1) {
				throw new UnsupportedOperationException("仅允许定义color/backColor/align/valign ！");
			}
			if (++verifyCount > 1) {
				throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
			}
			if (msgInfo.length > 0) {
				this.verifyIntNum = verifyIntNum + "@" + msgInfo[0];
			} else {
				this.verifyIntNum = verifyIntNum;
			}
			return this;
		}

		/**
		 * 浮点数字数据校验
		 * 
		 * @param verifyNum 表达式,请填写例如10.0~30.0
		 * @param msgInfo   提示消息
		 * @return
		 */
		public Column verifyFloatNum(String verifyFloatNum, String... msgInfo) {
			if (style == 1) {
				throw new UnsupportedOperationException("仅允许定义color/backColor/align/valign ！");
			}
			if (++verifyCount > 1) {
				throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
			}
			if (msgInfo.length > 0) {
				this.verifyFloatNum = verifyFloatNum + "@" + msgInfo[0];
			} else {
				this.verifyFloatNum = verifyFloatNum;
			}
			return this;
		}

		String getVerifyText() {
			return verifyText;
		}

		String getVerifyCustom() {
			return verifyCustom;
		}

		/**
		 * 自定义表达式校验
		 * 
		 * @param verifyCustom 表达式 ， 注意！！！xls格式和xlsx格式的表达式不太一样，xls从当前位置A1开始算起，xlsx从当前位置开始算起,已经修正过了
		 * @param msgInfo      提示消息
		 * @return
		 */
		public Column verifyCustom(String verifyCustom, String... msgInfo) {
			if (style == 1) {
				throw new UnsupportedOperationException("仅允许定义color/backColor/align/valign ！");
			}
			if (++verifyCount > 1) {
				throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
			}
			if (msgInfo.length > 0) {
				this.verifyCustom = verifyCustom + "@" + msgInfo[0];
			} else {
				this.verifyCustom = verifyCustom;
			}
			return this;
		}

		String getVerifyFloatNum() {
			return verifyFloatNum;
		}

		/**
		 * 文本长度校验
		 * 
		 * @param verifyText 比如输入1~2
		 * @param msgInfo    提示消息
		 * @return
		 */
		public Column verifyText(String verifyText, String... msgInfo) {
			if (style == 1) {
				throw new UnsupportedOperationException("仅允许定义color/backColor/align/valign ！");
			}
			if (++verifyCount > 1) {
				throw new UnsupportedOperationException("同一列只能定义一个数据校验！");
			}
			if (msgInfo.length > 0) {
				this.verifyText = verifyText + "@" + msgInfo[0];
			} else {
				this.verifyText = verifyText;
			}

			return this;
		}

	}

}
