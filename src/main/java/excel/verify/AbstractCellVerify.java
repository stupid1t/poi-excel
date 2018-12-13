package excel.verify;

/**
 * 列校验和格式化接口
 * 
 * @author 625
 *
 */
public abstract class AbstractCellVerify {
	public abstract Object verify(Object cellValue) throws Exception;
}
