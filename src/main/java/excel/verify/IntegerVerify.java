package excel.verify;

import org.apache.commons.lang3.StringUtils;

import excel.POIException;

/**
 * int数据校验
 *
 * @author 625
 */
public class IntegerVerify extends AbstractCellVerify {
	private String cellName;
	private AbstractCellValueVerify cellValueVerify;
	private boolean allowNull;

	public IntegerVerify(String cellName, boolean allowNull) {
		this.cellName = cellName;
		this.allowNull = allowNull;
	}

	public IntegerVerify(String cellName, AbstractCellValueVerify cellValueVerify, boolean allowNull) {
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

		cellValue = format(cellValue);
		if (null != cellValueVerify) {
			cellValue = cellValueVerify.verify(cellValue);
		}
		return cellValue;
	}

	private int format(Object fileValue) {
		int value;
		try {
			value = Double.valueOf(String.valueOf(fileValue)).intValue();
		} catch (Exception e) {
			throw POIException.newMessageException(cellName + "格式不正确:" + fileValue);
		}
		return value;
	}
}
