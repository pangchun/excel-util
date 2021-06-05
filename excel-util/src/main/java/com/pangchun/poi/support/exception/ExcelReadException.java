package com.pangchun.poi.support.exception;

/**
 * @author pangchun
 * @since 2021/6/5
 * @description excel读入异常
 */
public class ExcelReadException extends RuntimeException{

    public ExcelReadException(String message) {
        super(message);
    }
}