package com.github.stupdit1t.excel.common;

import java.io.Serializable;
import java.util.function.Function;

public interface Fn<T, R> extends Function<T, R>, Serializable {
}
