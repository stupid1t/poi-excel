package com.github.stupdit1t.excel.common;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * excel 导入返回的实体类
 *
 * @param <T>
 * @author 625
 */
public class PoiResult<T> {

    private boolean hasError = false;

    private List<ErrorMessage> error = new ArrayList<>();

    private List<T> data = Collections.emptyList();

    /**
     * 获取数据
     *
     * @return
     */
    public List<T> getData() {
        return data;
    }


    /**
     * 获取解析产生的错误
     *
     * @return
     */
    public List<ErrorMessage> getError() {
        return error;
    }

    /**
     * 判断是否有错误
     *
     * @return
     */
    public boolean hasError() {
        return hasError;
    }

    /**
     * 判断是否有数据产生
     */
    public boolean hasData() {
        return this.data != null && !this.data.isEmpty();
    }

    /**
     * 获取POIException 格式化失败信息
     *
     * @return
     */
    public List<String> getErrorInfo() {
        if (this.hasError()) {
            return error.stream().map(e -> {
                String location = e.getLocation();
                Exception exception = e.getException();
                return String.format(PoiConstant.ROW_INDEX_STR, location, exception.getMessage());
            }).collect(Collectors.toList());
        }
        return Collections.emptyList();
    }

    /**
     * 获取行级别 错误
     *
     * @return
     */
    public List<String> getErrorInfoLine() {
        if (this.hasError()) {
            List<String> message = new ArrayList<>();
            Map<Integer, List<ErrorMessage>> lineByError = error.stream().collect(Collectors.groupingBy(ErrorMessage::getRow, Collectors.toList()));
            lineByError.forEach((rowNum, errorMessage) -> {
                List<String> subMessage = errorMessage.stream().map(e -> {
                    int col = e.getCol();
                    int row = e.getRow();
                    String location = e.getLocation();
                    Exception exception = e.getException();
                    if (col != -1 && row != -1) {
                        return String.format("%s-%s", location, exception.getMessage());
                    } else {
                        return exception.getMessage();
                    }
                }).collect(Collectors.toList());
                if (rowNum != -1) {
                    message.add(String.format(PoiConstant.ROW_INDEX_STR, "第" + (rowNum + 1) + "行", String.join(" ", subMessage)));
                } else {
                    message.add(String.format(PoiConstant.ROW_INDEX_STR, "未知", String.join(" ", subMessage)));
                }
            });
            return message;
        }
        return Collections.emptyList();
    }

    /**
     * 获取行级别 字符串错误
     *
     * @return
     */
    public String getErrorInfoLineString() {
        return getErrorInfoLineString("\n");
    }


    /**
     * 获取行级别 字符串错误
     *
     * @param delimiter 每个错误的分隔符
     * @return
     */
    public String getErrorInfoLineString(String delimiter) {
        return String.join(delimiter, getErrorInfoLine());
    }


    /**
     * 获取字符串错误
     *
     * @return
     */
    public String getErrorInfoString() {
        return getErrorInfoString("\n");
    }

    /**
     * 获取字符串错误
     *
     * @param delimiter 每个错误的分隔符
     * @return
     */
    public String getErrorInfoString(String delimiter) {
        return String.join(delimiter, getErrorInfo());
    }

    public void setError(List<ErrorMessage> error) {
        this.error = error;
    }

    public static <T> PoiResult<T> fail(ErrorMessage e) {
        PoiResult<T> poiResult = new PoiResult<>();
        if (e != null) {
            poiResult.setHasError(true);
            poiResult.setError(Collections.singletonList(e));
        }
        return poiResult;
    }

    public void setData(List<T> beans) {
        this.data = beans;
    }

    public void setHasError(boolean hasError) {
        this.hasError = hasError;
}
}

