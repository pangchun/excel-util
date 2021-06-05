package com.pangchun.poi.support.bean;

import lombok.Data;
import lombok.experimental.Accessors;

@Data
@Accessors(chain = true)
public class ImageBean {

    /* 图片位置开始行号 */
    private Integer firstRow;

    /* 图片位置结束行号 */
    private Integer lastRow;

    /* 图片位置开始列号 */
    private Integer firstCol;

    /* 图片位置结束列号 */
    private Integer lastCol;

    /* 图片访问路径 */
    private String url;
}