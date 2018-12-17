package excel.verify;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 公共抽象校验类
 * 
 * @author 625
 *
 */
public abstract class AbstractVerifyBuidler {

	/**
	 * 字段校验集
	 */
	protected List<CellVerifyEntity> cellEntitys = new ArrayList<>();
	/**
	 * 字段名称
	 */
	public String[] filedNames;
	/**
	 * key:cellName, value:对应的校验类
	 */
	private Map<String, AbstractCellVerify> verifys;
	/**
	 * 列坐标
	 */
	public String[] cellRefs;

	/**
	 * 初始化
	 */
	public void init() {
		// 1、初始化filedNames
		filedNames = new String[cellEntitys.size()];
		// 2、初始化cellRefs
		cellRefs = new String[cellEntitys.size()];
		// 3、初始化verifys
		verifys = new HashMap<>(cellEntitys.size());
		for (int i = 0; i < cellEntitys.size(); i++) {
			CellVerifyEntity item = cellEntitys.get(i);
			verifys.put(item.getCellName(), item.getCellVerify());
			cellRefs[i] = item.getCellRef();
			filedNames[i] = item.getCellName();
		}
	}

	public Object verify(String fileName, Object fileValue) throws Exception {
		if (verifys == null) {
			throw new Exception("AbstractVerifyBuidler的子类需要调用父类的init进行初始化！");
		}
		if (verifys.containsKey(fileName)) {
			AbstractCellVerify verify = verifys.get(fileName);
			return verify.verify(fileValue);
		}
		return fileValue;
	}
	
	public Map<String, AbstractCellVerify> getVerifys() {
		return verifys;
	}


}
