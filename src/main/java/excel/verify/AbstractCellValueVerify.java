package excel.verify;

/**
 * 列值校验
 *
 */
public abstract class AbstractCellValueVerify {
	public abstract Object verify(Object fileValue) throws Exception;
}
