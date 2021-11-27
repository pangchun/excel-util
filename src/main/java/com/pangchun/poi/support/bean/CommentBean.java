package com.pangchun.poi.support.bean;

/**
 * @author pangchun
 * @since 2021/6/15
 * @description 批注信息，用于为检测到的不规范输入添加批注，以便用户修改
 */
public class CommentBean {

    /** 开始行 */
    private int firstRow;

    /** 结束行 */
    private int lastRow;

    /** 开始列 */
    private int firstCol;

    /** 结束列 */
    private int lastCol;

    /** 批注文本 */
    private String message;

    public CommentBean() {
    }

    public CommentBean(int firstRow, int lastRow, int firstCol, int lastCol, String message) {
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstCol = firstCol;
        this.lastCol = lastCol;
        this.message = message;
    }

    public int getFirstRow() {
        return firstRow;
    }

    public void setFirstRow(int firstRow) {
        this.firstRow = firstRow;
    }

    public int getLastRow() {
        return lastRow;
    }

    public void setLastRow(int lastRow) {
        this.lastRow = lastRow;
    }

    public int getFirstCol() {
        return firstCol;
    }

    public void setFirstCol(int firstCol) {
        this.firstCol = firstCol;
    }

    public int getLastCol() {
        return lastCol;
    }

    public void setLastCol(int lastCol) {
        this.lastCol = lastCol;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }
}