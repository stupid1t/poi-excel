package com.github.stupdit1t.excel.core;

/**
 * 链式钩子父 模式声明
 *
 * @param <T>
 */
public abstract class AbsParent<T> {

    /**
     * 上一个
     */
    public T parent;

    /**
     * 声明构造
     *
     * @param parent 当前对象
     */
    public AbsParent(T parent) {
        this.parent = parent;
    }

    /**
     * 返回父级
     *
     * @return T
     */
    public T done() {
        return parent;
    }
}
