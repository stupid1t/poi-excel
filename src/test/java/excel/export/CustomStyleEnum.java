package com.github.stupdit1t.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.function.BiConsumer;

/**
 * 默认单元格样式定制
 *
 * @author: 李涛
 * @version: 2020年09月03日 10:57
 */
public enum CustomStyleEnum implements ICellStyle {

	/**
	 * 标题样式
	 */
	TITLE(CellPosition.TITLE, (font, style) -> {
		font.setFontHeightInPoints((short) 15);
		font.setColor(IndexedColors.RED.getIndex());
		font.setBold(true);
		// 左右居中
		style.setAlignment(HorizontalAlignment.CENTER);
		// 上下居中
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setFont(font);
	}),
	;

	/**
	 * 位置
	 */
	private CellPosition position;

	/**
	 * 处理样式
	 */
	private BiConsumer<Font, CellStyle> customizeStyle;

	CustomStyleEnum(CellPosition position, BiConsumer<Font, CellStyle> customizeStyle) {
		this.position = position;
		this.customizeStyle = customizeStyle;
	}

	@Override
	public CellPosition getPosition() {
		return this.position;
	}

	@Override
	public void handleStyle(Font font, CellStyle cellStyle) {
		this.customizeStyle.accept(font, cellStyle);
	}

}
