package com.github.stupdit1t.excel.common;

/**
 * 异常定义
 *
 * @author 625
 */
public class PoiException extends RuntimeException {

    public PoiException(String message) {
        super(message);
    }

    public static PoiException error(String message) {
        return new PoiException(message);
    }


}
