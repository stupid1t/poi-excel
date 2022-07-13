package com.github.stupdit1t.excel.core.export;

import com.github.stupdit1t.excel.common.PoiCommon;
import com.github.stupdit1t.excel.core.AbsParent;
import com.github.stupdit1t.excel.style.CellPosition;

import java.util.ArrayList;
import java.util.List;

/**
 * 导出规则定义
 */
public class OpsSheet<R> extends AbsParent<OpsExport> {

	/**
	 * 标题高度
	 */
	short titleHeight = -1;

	/**
	 * 表头高度
	 */
	short headerHeight = -1;

	/**
	 * 单元格高度
	 */
	short cellHeight = -1;

	/**
	 * 尾行高度
	 */
	short footerHeight = -1;

	/**
	 * 列宽自定义
	 */
	int width = -1;

	/**
	 * 是否自动带序号
	 */
	boolean autoNum;

	/**
	 * 自动排序列宽度
	 */
	int autoNumColumnWidth = -1;

	/**
	 * 自定义合并的单元格
	 */
	List<Integer[]> mergerCells;

	/**
	 * sheet名字
	 */
	String sheetName;

	/**
	 * 导出的数据
	 */
	List<R> data;

	/**
	 * 导出的表头定义
	 */
	OpsHeader<R> opsHeader;

	/**
	 * 导出的数据列定义
	 */
	OpsColumn<R> opsColumn;

	/**
	 * 复杂尾设计容器
	 */
	OpsFooter<R> opsFooter;

	OpsSheet(OpsExport opsExport) {
		super(opsExport);
	}

	/**
	 * 表头设置
	 *
	 * @return OpsHeader<R>
	 */
	public OpsHeader<R> opsHeader() {
		this.opsHeader = new OpsHeader<>(this);
		return this.opsHeader;
	}

	/**
	 * 数据列定义
	 *
	 * @return OpsColumn<R>
	 */
	public OpsColumn<R> opsColumn() {
		this.opsColumn = new OpsColumn<>(this);
		return this.opsColumn;
	}

	/**
	 * 表头设置
	 *
	 * @return OpsSheet<R>
	 */
	public OpsFooter<R> opsFooter() {
		this.opsFooter = new OpsFooter<>(this);
		return this.opsFooter;
	}

	/**
	 * sheetName 定义
	 *
	 * @param sheetName sheet名称
	 * @return OpsSheet<R>
	 */
	public OpsSheet<R> sheetName(String sheetName) {
		this.sheetName = sheetName;
		return this;
	}

	/**
	 * sheetName 定义
	 * 自动生成序号, 此功能在复杂表头下, 需要自己定义序号列表头
	 *
	 * @return OpsSheet<R>
	 */
	public OpsSheet<R> autoNum() {
		this.autoNum = true;
		return this;
	}

	/**
	 * 自动序号列宽度
	 *
	 * @param autoNumColumnWidth 默认2000
	 * @return OpsSheet<R>
	 */
	public OpsSheet<R> autoNumColumnWidth(int autoNumColumnWidth) {
		this.autoNumColumnWidth = autoNumColumnWidth;
		return this;
	}

	/**
	 * 全局高度定义
	 *
	 * @param cellPosition 单元格类型
	 * @param height       高度
	 * @return OpsSheet<R>
	 */
	public OpsSheet<R> height(CellPosition cellPosition, int height) {
		switch (cellPosition) {
			case FOOTER:
				this.footerHeight = (short) height;
				break;
			case CELL:
				this.cellHeight = (short) height;
				break;
			case TITLE:
				this.titleHeight = (short) height;
				break;
			case HEADER:
				this.headerHeight = (short) height;
				break;
		}
		return this;
	}

	/**
	 *  全局列宽自定义
	 *
	 * @param width       宽度
	 * @return OpsSheet<R>
	 */
	public OpsSheet<R> width(int width) {
		this.width = width;
		return this;
	}

	/**
	 * 合并单元格
	 *
	 * @param location 坐标 A1:B2 或 1,2,A,B 这样
	 * @return OpsSheet<R>
	 */
	public OpsSheet<R> mergeCell(String location) {
		return mergeCell(PoiCommon.coverRangeIndex(location));
	}

	/**
	 * 合并单元格
	 *
	 * @param locationIndex 数组下标 如 0,0,0,0
	 * @return OpsSheet<R>
	 */
	public OpsSheet<R> mergeCell(Integer... locationIndex) {
		if (mergerCells == null) {
			mergerCells = new ArrayList<>();
		}
		mergerCells.add(locationIndex);
		return this;
	}

	/**
	 * 合并单元格
	 *
	 * @param locations 批量合并坐标
	 * @return OpsSheet<R>
	 */
	public OpsSheet<R> mergeCells(List<String> locations) {
		for (String location : locations) {
			mergeCell(location);
		}
		return this;
	}

	/**
	 * 合并单元格
	 *
	 * @param locations 批量合并坐标
	 * @return OpsSheet<R>
	 */
	public OpsSheet<R> mergeCellsIndex(List<Integer[]> locations) {
		for (Integer[] location : locations) {
			mergeCell(location);
		}
		return this;
	}
}