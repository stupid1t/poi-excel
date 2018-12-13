package excel.verify;

/**
 * 列校验实体
 * 
 * @author Administrator
 *
 */
public class CellVerifyEntity {

	/**
	 * 列名称 是否必要： 必要 作用：用来绑定对象属性
	 */
	private String cellName;
	/**
	 * 列坐标 是否必要： 必要 作用： 用来提取对应列
	 */
	private String cellRef;
	/**
	 * 列校验 是否必要： 非必要 作用：用来校验列值
	 */
	private AbstractCellVerify cellVerify;

	public CellVerifyEntity() {
		super();
	}

	public CellVerifyEntity(String cellName, String cellRef) {
		super();
		this.cellName = cellName;
		this.cellRef = cellRef;
	}

	public CellVerifyEntity(String cellName, String cellRef, AbstractCellVerify cellVerify) {
		super();
		this.cellName = cellName;
		this.cellRef = cellRef;
		this.cellVerify = cellVerify;
	}

	public String getCellName() {
		return cellName;
	}

	public void setCellName(String cellName) {
		this.cellName = cellName;
	}

	public AbstractCellVerify getCellVerify() {
		return cellVerify;
	}

	public void setCellVerify(AbstractCellVerify cellVerify) {
		this.cellVerify = cellVerify;
	}

	public String getCellRef() {
		return cellRef;
	}

	public void setCellRef(String cellRef) {
		this.cellRef = cellRef;
	}
}
