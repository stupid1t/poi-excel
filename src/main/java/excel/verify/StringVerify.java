package excel.verify;

import org.apache.commons.lang3.StringUtils;

import excel.POIException;


/**
 * 字符值校验实体
 * 
 * @author 625
 *
 */
public class StringVerify extends AbstractCellVerify {
	private String cellName;
	private AbstractCellValueVerify cellValueVerify;
	private boolean allowNull;

	public StringVerify(String cellName, boolean allowNull) {
		this.cellName = cellName;
		this.allowNull = allowNull;
	}

	public StringVerify(String cellName, AbstractCellValueVerify cellValueVerify, boolean allowNull) {
		super();
		this.cellName = cellName;
		this.cellValueVerify = cellValueVerify;
		this.allowNull = allowNull;
	}

	@Override
	public Object verify(Object cellValue) throws Exception {
		if (allowNull) {
			if (cellValue != null && StringUtils.isNotEmpty(format(cellValue))) {
				cellValue = format(cellValue);
				if (null != cellValueVerify) {
					cellValue = cellValueVerify.verify(cellValue);
				}
				return cellValue;
			}
			return cellValue;
		}

		if (cellValue == null || StringUtils.isEmpty(format(cellValue))) {
			throw POIException.newMessageException(cellName + "不能为空");
		}

		cellValue = format(cellValue);
		if (null != cellValueVerify) {
			cellValue = cellValueVerify.verify(cellValue);
		}
		return cellValue;
	}

	private String format(Object fileValue) {
		return String.valueOf(fileValue);
	}
}
