package com.github.stupdit1t.excel.core.export;

import com.github.stupdit1t.excel.common.PoiCommon;
import com.github.stupdit1t.excel.core.AbsParent;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;

public class OpsFooter<R> extends AbsParent<OpsSheet<R>> {

	/**
	 * 复杂表头设计容器
	 */
	List<ComplexCell> complexFooter;

	/**
	 * 声明headers
	 *
	 * @param export parent
	 */
	OpsFooter(OpsSheet<R> export) {
		super(export);
	}

	/**
	 * 获取复杂表头设计
	 *
	 * @param text     表头文字
	 * @param location 文字位置, 如填写1,3,A,C 或 A1:C3
	 * @return List<ComplexHeader < R>>
	 */
	public OpsFooter<R> text(String text, String location) {
		return this.text(text, PoiCommon.coverRangeIndex(location));
	}

	/**
	 * 获取复杂表头设计
	 *
	 * @param text     表头文字
	 * @param location 文字位置, 如填写1,3,A,C 或 A1:C3
	 * @param style    表头样式
	 * @return List<ComplexHeader < R>>
	 */
	public OpsFooter<R> text(String text, String location, BiConsumer<Font, CellStyle> style) {
		this.text(text, PoiCommon.coverRangeIndex(location), style);
		return this;
	}

	/**
	 * 获取复杂表头设计
	 *
	 * @param text          表头位置
	 * @param locationIndex 表头位置, 下标0开始, 如 0,2,0,2
	 * @return List<ComplexHeader < R>>
	 */
	public OpsFooter<R> text(String text, Integer... locationIndex) {
		return text(text, locationIndex, null);
	}

	/**
	 * 获取复杂表头设计
	 *
	 * @param text          表头位置
	 * @param locationIndex 表头位置, 下标0开始, 如 0,2,0,2
	 * @param style         表头样式
	 * @return List<ComplexHeader < R>>
	 */
	public OpsFooter<R> text(String text, Integer[] locationIndex, BiConsumer<Font, CellStyle> style) {
		if (complexFooter == null) {
			complexFooter = new ArrayList<>();
		}
		ComplexCell complexCell = new ComplexCell();
		complexCell.setText(text);
		complexCell.setLocationIndex(locationIndex);
		complexCell.setStyle(style);
		complexFooter.add(complexCell);
		return this;
	}
}
