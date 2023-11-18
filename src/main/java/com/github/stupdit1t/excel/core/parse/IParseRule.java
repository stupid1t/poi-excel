package com.github.stupdit1t.excel.core.parse;

import java.util.function.Function;

public interface IParseRule<R> {

    /**
     * 不能为空
     *
     * @return InColumn<R>
     */
    IParseRule<R> notNull();


    /**
     * 去除两边空格
     *
     * @return InColumn<R>
     */
    IParseRule<R> trim();

    /**
     * 去除两边空格
     *
     * @return InColumn<R>
     */
    IParseRule<R> defaultValue(Object defaultValue);

    /**
     * 正则校验
     * @param regex
     * @return
     */
    IParseRule<R> regex(String regex);

    /**
     * 格式化，日期
     * @param format
     * @return
     */
    IParseRule<R> format(String format);

    /**
     * 如果是数字设置精度
     *
     * @param precision
     * @return
     */
    IParseRule<R> scale(int precision);



    /**
     * 单元格值转换处理，验证
     *
     * @param mapping
     * @return
     */
    IParseRule<R> map(Function<Object, Object> mapping);

    /**
     * 如果转map，进行类型强制转换，非必须
     *
     * @param covertCls
     * @return
     */
    IParseRule<R> type(Class<?> covertCls);


    OpsSheet<R> done();


    IParseRule<R> field(String index, String field);
}
