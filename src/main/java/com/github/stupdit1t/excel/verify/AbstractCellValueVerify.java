package com.github.stupdit1t.excel.verify;

/**
 * 列值校验
 *
 * @author 625
 */
public abstract class AbstractCellValueVerify {

	/**
	 * 校验单元格值
	 *
	 * @param fileValue
	 * @return Object
	 * @throws Exception
	 */
	public abstract Object verify(Object fileValue) throws Exception;
}
