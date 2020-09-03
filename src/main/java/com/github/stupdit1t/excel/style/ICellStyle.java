package com.github.stupdit1t.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

/**
 * 单元格样式定义
 *
 */
public interface ICellStyle {

	/**
	 * 样式位置
	 */
	CellPosition getPosition();


	/**
	 * 样式处理方式
	 *
	 * @param font      当前字体
	 * @param cellStyle 当前单元格样式
	 */
	void handleStyle(Font font, CellStyle cellStyle);

	/**
	 * 根据位置获取样式
	 *
	 * @param position   位置
	 * @param cellStyles 样式
	 */
	static ICellStyle getCellStyleByPosition(CellPosition position, ICellStyle[] cellStyles) {
		for (ICellStyle cellStyle : cellStyles) {
			if (cellStyle.getPosition() == position) {
				return cellStyle;
			}
		}

		// 找不到取默认值
		ICellStyle[] defaultCellStyle = DefaultCellStyleEnum.values();
		for (ICellStyle cellStyle : defaultCellStyle) {
			if (cellStyle.getPosition() == position) {
				return cellStyle;
			}
		}
		throw new UnsupportedOperationException("找不到对应的样式 " + position);
	}
}
