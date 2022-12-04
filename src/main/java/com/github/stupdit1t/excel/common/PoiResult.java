package com.github.stupdit1t.excel.common;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * excel 导入返回的实体类
 *
 * @param <T>
 * @author 625
 */
public class PoiResult<T> {

    private boolean success = true;

    private List<String> message = Collections.emptyList();

    private List<Exception> unknownError = new ArrayList<>();

    private List<T> data;

    public boolean isSuccess() {
        return success;
    }

    public void setSuccess(boolean success) {
        this.success = success;
    }

    public List<String> getMessage() {
        return message;
    }

    public String getMessageToString() {
        return String.join("\n", message);
    }

    public void setMessage(List<String> message) {
        this.message = message;
    }

    public List<T> getData() {
        return data;
    }

    public void setData(List<T> beans) {
        this.data = beans;
    }

    public List<Exception> getUnknownError() {
        return unknownError;
    }

    public void setUnknownError(List<Exception> unknownError) {
        this.unknownError = unknownError;
    }

    public static <T> PoiResult<T> fail() {
        PoiResult<T> poiResult = new PoiResult<>();
        poiResult.setSuccess(false);
        poiResult.setMessage(Collections.singletonList("读取Excel失败"));
        poiResult.setUnknownError(Collections.emptyList());
        poiResult.setData(Collections.emptyList());
        return poiResult;
    }

}
