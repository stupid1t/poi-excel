package com.github.stupdit1t.excel.common;

public class ErrorMessage {

    private String location;

    private int row;

    private int col;

    private Exception exception;

    public ErrorMessage(String location, int row, int col, Exception exception) {
        this.location = location;
        this.row = row;
        this.col = col;
        this.exception = exception;
    }

    public ErrorMessage(Exception exception) {
        this.row = -1;
        this.col = -1;
        this.exception = exception;
    }

    public String getLocation() {
        return location;
    }

    public int getRow() {
        return row;
    }

    public int getCol() {
        return col;
    }

    public Exception getException() {
        return exception;
    }
}
