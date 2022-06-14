package com.github.stupdit1t.excel.core.export;

import com.github.stupdit1t.excel.common.PoiCommon;
import com.github.stupdit1t.excel.core.AbsParent;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.function.BiConsumer;

public class OpsHeader<R> extends AbsParent<OpsSheet<R>> {

	/**
	 * 复杂表头设计容器
	 */
	ComplexHeader<R> complex;

	/**
	 * 简单表头设计容器
	 */
	SimpleHeader<R> simple;

	/**
	 * 是否冻结表头
	 */
	boolean freeze = true;

	/**
	 * 设置状态
	 * 0 未设置 1 复杂表头 2 简单表头
	 */
	int mode;

	/**
	 * 声明headers
	 *
	 * @param export parent
	 */
	OpsHeader(OpsSheet<R> export) {
		super(export);
	}

	/**
	 * 检测是否已经被设置状态
	 *
	 * @param wantSetStatus 要设置的状态
	 */
	private void checkSetStatus(int wantSetStatus) {
		if (mode != 0 && mode != wantSetStatus) {
			throw new UnsupportedOperationException("不支持 simple 表头和 complex 表头同时设置");
		}
		this.mode = wantSetStatus;
	}

	/**
	 * 冻结表头
	 *
	 * @param freeze 冻结表头
	 * @return OpsHeader<R>
	 */
	public OpsHeader<R> freeze(boolean freeze) {
		this.freeze = freeze;
		return this;
	}

	/**
	 * 简单表头构建
	 *
	 * @return SimpleHeader
	 */
	public SimpleHeader<R> simple() {
		checkSetStatus(1);
		this.simple = new SimpleHeader<>(this.parent);
		return simple;
	}

	/**
	 * 简单表头构建
	 *
	 * @return SimpleHeader
	 */
	public ComplexHeader<R> complex() {
		checkSetStatus(2);
		this.complex = new ComplexHeader<>(this.parent);
		return complex;
	}

	/**
	 * 简单表头定义
	 */
	public static class SimpleHeader<R> extends AbsParent<OpsSheet<R>> {

		/**
		 * 大标题
		 */
		String title;

		/**
		 * header文本设置
		 */
		LinkedHashMap<String, BiConsumer<Font, CellStyle>> headers = new LinkedHashMap<>();

		/**
		 * 声明
		 *
		 * @param opsSheet parent
		 */
		SimpleHeader(OpsSheet<R> opsSheet) {
			super(opsSheet);
		}

		/**
		 * 标题设置
		 *
		 * @param title 大标题
		 * @return SimpleHeader<R>
		 */
		public SimpleHeader<R> title(String title) {
			this.title = title;
			return this;
		}

		/**
		 * 表头设置
		 *
		 * @param texts 表头
		 * @return SimpleHeader<R>
		 */
		public SimpleHeader<R> texts(String... texts) {
			for (String text : texts) {
				this.headers.put(text, null);
			}
			return this;
		}

		/**
		 * 表头设置
		 *
		 * @param text  文本
		 * @param style 样式
		 * @return SimpleHeader<R>
		 */
		public SimpleHeader<R> text(String text, BiConsumer<Font, CellStyle> style) {
			this.headers.put(text, style);
			return this;
		}

	}

	/**
	 * 复杂表头定义
	 */
	public static class ComplexHeader<R> extends AbsParent<OpsSheet<R>> {

		/**
		 * 表头规则
		 */
		List<ComplexCell> headers;

		/**
		 * 声明
		 *
		 * @param opsSheet parent
		 */
		ComplexHeader(OpsSheet<R> opsSheet) {
			super(opsSheet);
			headers = new ArrayList<>();
		}

		/**
		 * 获取复杂表头设计
		 *
		 * @param text     显示文本
		 * @param location 文本定位 , 如填写1,3,A,C 或 A1:C3
		 * @return List<ComplexHeader < R>>
		 */
		public ComplexHeader<R> text(String text, String location) {
			return this.text(text, PoiCommon.coverRangeIndex(location));
		}

		/**
		 * 获取复杂表头设计
		 *
		 * @param text     显示文本
		 * @param location 文本定位 , 如填写1,3,A,C 或 A1:C3
		 * @param style    显示样式
		 * @return List<ComplexHeader < R>>
		 */
		public ComplexHeader<R> text(String text, String location, BiConsumer<Font, CellStyle> style) {
			return this.text(text, PoiCommon.coverRangeIndex(location), style);
		}

		/**
		 * 获取复杂表头设计
		 *
		 * @param text          显示文本
		 * @param locationIndex 文本定位 , 如填写0,2,0,2
		 * @return List<ComplexHeader < R>>
		 */
		public ComplexHeader<R> text(String text, Integer... locationIndex) {
			return text(text, locationIndex, null);
		}

		/**
		 * 获取复杂表头设计
		 *
		 * @param text          显示文本
		 * @param locationIndex 文本定位 , 如填写0,2,0,2
		 * @param style         样式
		 * @return List<ComplexHeader < R>>
		 */
		public ComplexHeader<R> text(String text, Integer[] locationIndex, BiConsumer<Font, CellStyle> style) {
			ComplexCell complexCell = new ComplexCell();
			complexCell.setText(text);
			complexCell.setLocationIndex(locationIndex);
			complexCell.setStyle(style);
			headers.add(complexCell);
			return this;
		}

	}

}
