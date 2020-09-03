package com.github.stupdit1t.excel.verify;

import com.github.stupdit1t.excel.common.POIException;
import org.apache.commons.lang3.StringUtils;



/**
 * long校验实体
 * 
 * @author 625
 *
 */
public class LongVerify extends AbstractCellVerify {
	private String cellName;
	private AbstractCellValueVerify cellValueVerify;
	private boolean allowNull;

	public LongVerify(String cellName, boolean allowNull) {
		this.cellName = cellName;
		this.allowNull = allowNull;
	}

	public LongVerify(String cellName, AbstractCellValueVerify cellValueVerify, boolean allowNull) {
		super();
		this.cellName = cellName;
		this.cellValueVerify = cellValueVerify;
		this.allowNull = allowNull;
	}

	@Override
	public Object verify(Object cellValue) throws Exception {
		if (allowNull) {
			if (cellValue != null && StringUtils.isNotEmpty(String.valueOf(cellValue))) {
				cellValue = format(cellValue);
				if (null != cellValueVerify) {
					cellValue = cellValueVerify.verify(cellValue);
				}
				return cellValue;
			}
			return cellValue;
		}

		if (cellValue == null || StringUtils.isEmpty(String.valueOf(cellValue))) {
			throw POIException.newMessageException(cellName + "不能为空");
		}

		if (format(cellValue) < 0) {
			throw POIException.newMessageException(cellName + "必须大于零");
		}

		cellValue = format(cellValue);
		if (null != cellValueVerify) {
			cellValue = cellValueVerify.verify(cellValue);
		}
		return cellValue;
	}

	private long format(Object fileValue) {

		long value;
		try {
			value = Long.valueOf(String.valueOf(fileValue).substring(0, String.valueOf(fileValue).indexOf(".")));
		} catch (Exception e) {
			throw POIException.newMessageException(cellName + "格式不正确:" + fileValue);
		}
		return value;
	}
}
