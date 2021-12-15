/*
 * Copyright (c) 2015-2018 SHENZHEN TOMTOP SCIENCE AND TECHNOLOGY DEVELOP CO., LTD. All rights reserved.
 *
 * 注意：本内容仅限于深圳市通拓科技研发有限公司内部传阅，禁止外泄以及用于其他的商业目的
 */
package excel.imports;

import com.github.stupdit1t.excel.verify.*;
import com.github.stupdit1t.excel.verify.rule.AbsSheetVerifyRule;
import com.github.stupdit1t.excel.verify.rule.CellVerifyRule;

import java.util.List;

/**
 * 导入用户校验类
 *
 * @author 625
 */
public class ProjectVerifyBuilder extends AbsSheetVerifyRule {


    @Override
    protected void buildRule(List<CellVerifyRule> cellEntitys) {
        cellEntitys.add(new CellVerifyRule("B", "projectName", "项目名称", new StringVerify(true)));
        cellEntitys.add(new CellVerifyRule("C", "areaName", "所属区域", new StringVerify(true)));
        cellEntitys.add(new CellVerifyRule("D", "province", "省份", new StringVerify(true)));
        cellEntitys.add(new CellVerifyRule("E", "city", "市", new StringVerify(true)));
        cellEntitys.add(new CellVerifyRule("F", "people", "项目所属人", new StringVerify(false)));
        cellEntitys.add(new CellVerifyRule("G", "leader", "项目领导人", new StringVerify(true)));
        cellEntitys.add(new CellVerifyRule("H", "scount", "总分", new IntegerVerify(true)));
        cellEntitys.add(new CellVerifyRule("I", "avg", "历史平均分", new DoubleVerify(true)));
        cellEntitys.add(new CellVerifyRule("J", "createTime", "创建时间", new DateTimeVerify("yyyy-MM-dd dd:mm", true)));
        cellEntitys.add(new CellVerifyRule("K", "img", "图片", new ImgVerify(false)));
    }
}
