package com.github.stupdit1t.excel.style;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.function.BiConsumer;

/**
 * 默认单元格样式定制
 *
 */
public enum DefaultCellStyleEnum  implements ICellStyle {

	/**
	 * 标题样式
	 */
	TITLE(CellPosition.TITLE, (font, style) -> {
		font.setFontHeightInPoints((short) 15);
		font.setBold(true);
		// 左右居中
		style.setAlignment(HorizontalAlignment.CENTER);
		// 上下居中
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setFont(font);
	}),

	/**
	 * 副标题样式
	 */
	HEADER(CellPosition.HEADER, (font, style) -> {
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 10);
		font.setColor(IndexedColors.WHITE.getIndex());
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
		style.setFont(font);
		style.setWrapText(true);
	}),

	/**
	 * 单元格样式
	 */
	CELL(CellPosition.CELL, (font, style) -> {
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 10);
		style.setFont(font);
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
	});

	/**
	 * 位置
	 */
	private CellPosition position;

	/**
	 * 处理样式
	 */
	private BiConsumer<Font, CellStyle> customizeStyle;

	DefaultCellStyleEnum(CellPosition position, BiConsumer<Font, CellStyle> customizeStyle) {
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
