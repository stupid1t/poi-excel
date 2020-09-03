package com.github.stupdit1t.excel.verify;


import com.github.stupdit1t.excel.common.POIException;

/**
 * 字符值校验实体
 * 
 * @author 625
 *
 */
public class ImgVerify extends AbstractCellVerify {
	private String cellName;
	private AbstractCellValueVerify cellValueVerify;
	private boolean allowNull;

	public ImgVerify(String cellName, boolean allowNull) {
		this.cellName = cellName;
		this.allowNull = allowNull;
	}

	public ImgVerify(String cellName, AbstractCellValueVerify cellValueVerify, boolean allowNull) {
		super();
		this.cellName = cellName;
		this.cellValueVerify = cellValueVerify;
		this.allowNull = allowNull;
	}

	@Override
	public Object verify(Object cellValue) throws Exception {
		if (allowNull) {
			if (cellValue != null) {
				if (null != cellValueVerify) {
					cellValue = cellValueVerify.verify(cellValue);
				}
				return cellValue;
			}
			return cellValue;
		}
		if (cellValue == null) {
			throw POIException.newMessageException(cellName + "不能为空");
		}
		if (null != cellValueVerify) {
			cellValue = cellValueVerify.verify(cellValue);
		}
		return cellValue;
	}

}
