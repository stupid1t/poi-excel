package com.github.stupdit1t.excel.core.export;

import com.github.stupdit1t.excel.common.PoiWorkbookType;
import com.github.stupdit1t.excel.core.OpsPoiUtil;
import com.github.stupdit1t.excel.style.DefaultCellStyleEnum;
import com.github.stupdit1t.excel.style.ICellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.function.BiConsumer;

/**
 * 导出规则定义
 */
public class OpsExport implements OpsFinish {

	/**
	 * 输出的sheet
	 */
	List<OpsSheet<?>> opsSheets;

	/**
	 * 文件格式
	 */
	PoiWorkbookType workbookType;

	/**
	 * 当前工作簿
	 */
	Workbook workbook;

	/**
	 * 全局单元格样式
	 */
	ICellStyle[] style = DefaultCellStyleEnum.values();

	/**
	 * Excel密码, 只支持xls 格式
	 */
	String password;

	/**
	 * 并行导出sheet
	 */
	boolean parallelSheet = false;

	/**
	 * 输出模式
	 * 1 路径输出
	 * 2 流输出
	 * 3 servlet 响应
	 */
	int toMode;

	/**
	 * 输出目录
	 */
	String path;

	/**
	 * 输出流
	 */
	OutputStream stream;

	/**
	 * 输出 Servlet 响应
	 */
	HttpServletResponse response;

	/**
	 * 输出 Servlet 响应 文件名
	 */
	String responseName;

	public OpsExport(PoiWorkbookType workbookType) {
		this.workbookType = workbookType;
	}

	public OpsExport(Workbook workbook) {
		this.workbook = workbook;
		if( workbook instanceof SXSSFWorkbook){
			this.workbookType = PoiWorkbookType.BIG_XLSX;
		}else if(workbook instanceof XSSFWorkbook){
			this.workbookType = PoiWorkbookType.XLSX;
		}else if(workbook instanceof HSSFWorkbook){
			this.workbookType = PoiWorkbookType.XLS;
		}
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
	private void checkSetToMode(int wantSetMode) {
		if (toMode != 0 && toMode != wantSetMode) {
			throw new UnsupportedOperationException("仅支持设置输出 1 种输出方式");
		}
		this.toMode = wantSetMode;
	}

	/**
	 * 全局样式设置
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
	 * 并行导出sheet, 默认fork join线程池
	 *
	 * @return OpsExport
	 */
	public OpsExport parallelSheet() {
		this.parallelSheet = true;
		return this;
	}

	/**
	 * 输出路径设置
	 *
	 * @param toPath 输出磁盘路径
	 */
	@Override
	public void export(String toPath) {
		checkSetToMode(1);
		this.path = toPath;
		final Workbook workbook = this.getWorkBook();
		this.export(workbook);
	}

	/**
	 * 输出流
	 *
	 * @param toStream 输出流
	 */
	@Override
	public void export(OutputStream toStream) {
		checkSetToMode(2);
		this.stream = toStream;
		final Workbook workbook = this.getWorkBook();
		this.export(workbook);
	}

	/**
	 * 输出servlet
	 *
	 * @param toResponse 输出servlet
	 * @param fileName   文件名
	 */
	@Override
	public void export(HttpServletResponse toResponse, String fileName) {
		checkSetToMode(3);
		this.response = toResponse;
		this.responseName = fileName;
		final Workbook workbook = this.getWorkBook();
		this.export(workbook);
	}

	/**
	 * 执行输出
	 *
	 * @param workbook 导出workbook
	 */
	@Override
	public void export(Workbook workbook) {
		// 5.执行导出
		switch (this.toMode) {
			case 1:
				OpsPoiUtil.export(workbook, this.path, this.password);
				break;
			case 2:
				OpsPoiUtil.export(workbook, this.stream, this.password);
				break;
			case 3:
				OpsPoiUtil.export(workbook, this.response, this.responseName, this.password);
				break;
		}
	}

	/**
	 * 创建workbook
	 */
	public Workbook getWorkBook() {
		// 1.声明工作簿
		if(this.workbook == null){
			this.workbook =  workbookType.create();
		}

		// 2.多sheet获取
		if (this.parallelSheet) {
			CountDownLatch count = new CountDownLatch(opsSheets.size());
			opsSheets.parallelStream().forEach(opsSheet -> {
				fillBook(workbook, opsSheet);
				count.countDown();
			});
			try {
				count.await();
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		} else {
			for (OpsSheet<?> opsSheet : opsSheets) {
				fillBook(workbook, opsSheet);
			}
		}

		return workbook;
	}

	/**
	 * 填充book
	 *
	 * @param workbook 工作簿
	 * @param opsSheet sheet
	 */
	private void fillBook(Workbook workbook, OpsSheet<?> opsSheet) {
		// 必要参数检查
		checkHeader(opsSheet.opsHeader);
		checkColumn(opsSheet.opsColumn);
		if (opsSheet.data == null) {
			throw new UnsupportedOperationException("导出数据data不能设置为null!");
		}

		// header获取
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
		exportRules.setTitleHeight(opsSheet.titleHeight);
		exportRules.setHeaderHeight(opsSheet.headerHeight);
		exportRules.setCellHeight(opsSheet.cellHeight);
		exportRules.setFooterHeight(opsSheet.footerHeight);
		exportRules.setSheetName(opsSheet.sheetName);
		exportRules.setFreezeHeader(opsSheet.opsHeader.freeze);
		exportRules.setPassword(this.password);
		exportRules.setGlobalStyle(this.style);
		exportRules.setColumnWidth(opsSheet.width);
		if (opsSheet.autoNumColumnWidth != -1) {
			exportRules.setAutoNumColumnWidth(opsSheet.autoNumColumnWidth);
		}
		exportRules.setAutoNum(opsSheet.autoNum);
		exportRules.setMergerCells(opsSheet.mergerCells);
		exportRules.setImages(opsSheet.images);
		// footer内容提取
		if (opsSheet.opsFooter != null) {
			List<ComplexCell> footer = opsSheet.opsFooter.complexFooter;
			exportRules.setFooterRules(footer);
		}

		// 4.填充表格
		OpsPoiUtil.fillBook(workbook, opsSheet.data, exportRules);
	}


	/**
	 * 检测表头数据
	 *
	 * @param opsHeader 导出头
	 */
	private void checkHeader(OpsHeader<?> opsHeader) {
		if (opsHeader == null || (opsHeader.simple == null && opsHeader.complex == null)) {
			throw new UnsupportedOperationException("请至少设置一种导出方式, 如simple 或 complex");
		} else {
			// 复杂表头模式
			if (opsHeader.mode == 2) {
				List<ComplexCell> headers = opsHeader.complex.headers;
				if (headers == null || headers.isEmpty()) {
					throw new UnsupportedOperationException("请设置complex表头的text数据");
				}
			} else {
				OpsHeader.SimpleHeader<?> simple = opsHeader.simple;
				LinkedHashMap<String, BiConsumer<Font, CellStyle>> headers = simple.headers;
				if (headers == null || headers.isEmpty()) {
					throw new UnsupportedOperationException("请设置simple表头的text数据");
				}
			}
		}
	}

	/**
	 * 检测导出列设置
	 *
	 * @param opsColumn 导出头
	 */
	private void checkColumn(OpsColumn<?> opsColumn) {
		if (opsColumn == null || opsColumn.columns == null || opsColumn.columns.isEmpty()) {
			throw new UnsupportedOperationException("请设置opsColumn数据");
		}
	}
}