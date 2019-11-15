package excel.verify;

import excel.POIException;
import org.apache.commons.lang3.StringUtils;

import java.util.regex.Pattern;


/**
 * 正则校验
 * 
 * @author 625
 *
 */
public class RegExpVerify extends AbstractCellVerify {
	private String cellName;
	private String pattern;
	private AbstractCellValueVerify cellValueVerify;
	private boolean allowNull;

	public RegExpVerify(String cellName, String pattern, boolean allowNull) {
		this.cellName = cellName;
		this.pattern = pattern;
		this.allowNull = allowNull;
	}

	public RegExpVerify(String cellName, String pattern, AbstractCellValueVerify cellValueVerify, boolean allowNull) {
		super();
		this.cellName = cellName;
		this.pattern = pattern;
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

	private String format(Object fileValue) {
		String value = String.valueOf(fileValue);
		if (!Pattern.matches(pattern, value)) {
			throw POIException.newMessageException(cellName + "格式不正确:" + fileValue);
		}
		return value;
	}

}
