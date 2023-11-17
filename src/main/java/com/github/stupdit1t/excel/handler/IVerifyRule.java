package com.github.stupdit1t.excel.handler;

import com.github.stupdit1t.excel.core.AbsParent;
import com.github.stupdit1t.excel.core.parse.OpsColumn;
import org.apache.poi.ss.formula.functions.T;

import java.util.function.Function;

public interface IVerifyRule<R> {

    /**
     * 不能为空
     *
     * @return InColumn<R>
     */
    IVerifyRule<R> notNull();


    /**
     * 去除两边空格
     *
     * @return InColumn<R>
     */
    IVerifyRule<R> trim();

    /**
     * 去除两边空格
     *
     * @return InColumn<R>
     */
    IVerifyRule<R> defaultValue(Object defaultValue);

    /**
     * 正则校验
     * @param regex
     * @return
     */
    IVerifyRule<R> regex(String regex);

    /**
     * 格式化，日期
     * @param format
     * @return
     */
    IVerifyRule<R> format(String format);

    /**
     * 如果是数字设置精度
     *
     * @param precision
     * @return
     */
    IVerifyRule<R> scale(int precision);



    /**
     * 单元格值转换处理，验证
     *
     * @param mapping
     * @return
     */
    IVerifyRule<R> map(Function<Object, Object> mapping);

    /**
     * 如果转map，进行类型强制转换，非必须
     *
     * @param covertCls
     * @return
     */
    IVerifyRule<R> type(Class<?> covertCls);


    OpsColumn<R> done();

}
