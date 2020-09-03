/*
 * Copyright (c) 2015-2018 SHENZHEN TOMTOP SCIENCE AND TECHNOLOGY DEVELOP CO., LTD. All rights reserved.
 *
 * 注意：本内容仅限于深圳市通拓科技研发有限公司内部传阅，禁止外泄以及用于其他的商业目的 
 */
package excel.imports;

import com.github.stupdit1t.excel.verify.AbstractVerifyBuidler;
import com.github.stupdit1t.excel.verify.CellVerifyEntity;
import com.github.stupdit1t.excel.verify.DateTimeVerify;
import com.github.stupdit1t.excel.verify.StringVerify;

/**
 * 导入用户校验类
 * 
 * @author 625
 *
 */
public class EvaluateVerifyBuilder extends AbstractVerifyBuidler {

	private static EvaluateVerifyBuilder builder = new EvaluateVerifyBuilder();

	public static EvaluateVerifyBuilder getInstance() {
		return builder;
	}

	/**
	 * 定义列校验实体：提取的字段、提取列、校验规则
	 */
	private EvaluateVerifyBuilder() {
		cellEntitys.add(new CellVerifyEntity("projectName", "B", new StringVerify("项目名称", true)));
		cellEntitys.add(new CellVerifyEntity("areaName", "C", new StringVerify("所属区域", true)));
		cellEntitys.add(new CellVerifyEntity("province", "D", new StringVerify("省份", true)));
		cellEntitys.add(new CellVerifyEntity("city", "E", new StringVerify("市", true)));
		cellEntitys.add(new CellVerifyEntity("statusName", "F", new StringVerify("项目状态", true)));
		cellEntitys.add(new CellVerifyEntity("scount", "G", new StringVerify("总分", true)));
		cellEntitys.add(new CellVerifyEntity("areaInfo", "H", new StringVerify("区位条件", true)));
		cellEntitys.add(new CellVerifyEntity("resourceInfo", "I", new StringVerify("资源禀赋", true)));
		cellEntitys.add(new CellVerifyEntity("manageInfo", "G", new StringVerify("经营现状", true)));
		cellEntitys.add(new CellVerifyEntity("reviewInfo", "K", new StringVerify("考察印象", true)));
		cellEntitys.add(new CellVerifyEntity("teamInfo", "L", new StringVerify("管理团队", true)));
		cellEntitys.add(new CellVerifyEntity("img", "M", new StringVerify("风采", true)));
		cellEntitys.add(new CellVerifyEntity("createTime", "N", new DateTimeVerify("创建时间", "yyyy-MM-dd", true)));
		// 必须调用
		super.init();
	}
}
