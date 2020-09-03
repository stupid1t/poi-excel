package excel.imports;

import com.github.stupdit1t.excel.verify.AbstractVerifyBuidler;
import com.github.stupdit1t.excel.verify.CellVerifyEntity;
import com.github.stupdit1t.excel.verify.DoubleVerify;
import com.github.stupdit1t.excel.verify.ImgVerify;
import com.github.stupdit1t.excel.verify.IntegerVerify;
import com.github.stupdit1t.excel.verify.StringVerify;

public class ThingsVerifyBuilder extends AbstractVerifyBuidler {
	private static ThingsVerifyBuilder builder = new ThingsVerifyBuilder();

	public static ThingsVerifyBuilder getInstance() {
		return builder;
	}

	/**
	 * 定义列校验实体：提取的字段、提取列、校验规则
	 */
	private ThingsVerifyBuilder() {
		cellEntitys.add(new CellVerifyEntity("uuid", "B", new StringVerify("编码", false)));
		cellEntitys.add(new CellVerifyEntity("name", "C", new StringVerify("品名", false)));
		cellEntitys.add(new CellVerifyEntity("extraName", "D", new StringVerify("市面名称", true)));
		cellEntitys.add(new CellVerifyEntity("brand", "E", new StringVerify("品牌", true)));
		cellEntitys.add(new CellVerifyEntity("modelType", "F", new StringVerify("型号", true)));
		cellEntitys.add(new CellVerifyEntity("thingsTypeName", "G", new StringVerify("分类", false)));
		cellEntitys.add(new CellVerifyEntity("specifications", "H", new StringVerify("规格", true)));
		cellEntitys.add(new CellVerifyEntity("quality", "I", new StringVerify("材质", true)));
		cellEntitys.add(new CellVerifyEntity("pack", "J", new StringVerify("包装", true)));
		cellEntitys.add(new CellVerifyEntity("conversion", "K", new StringVerify("单位", false)));
		cellEntitys.add(new CellVerifyEntity("buyPrice", "L", new DoubleVerify("采购价", true)));
		cellEntitys.add(new CellVerifyEntity("marketPrice", "M", new DoubleVerify("销售价", true)));
		cellEntitys.add(new CellVerifyEntity("week", "N", new StringVerify("到货周期", true)));
		cellEntitys.add(new CellVerifyEntity("startNum", "O", new IntegerVerify("起订量", true)));
		cellEntitys.add(new CellVerifyEntity("pictureData", "P", new ImgVerify("图片", true)));
		cellEntitys.add(new CellVerifyEntity("thingsDesc", "Q", new StringVerify("描述", true)));

		// 必须调用
		super.init();
	}
}
