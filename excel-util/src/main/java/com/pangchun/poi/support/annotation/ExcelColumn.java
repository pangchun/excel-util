package com.pangchun.poi.support.annotation;

import java.lang.annotation.*;

/**
 * @author pangchun
 * @since 2021/6/5
 * @description 列索引注解，表示实体属性与单元格的列位置对应关系
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelColumn {

    /** 列索引，从0开始 */
    int index() default -1;

    /** 列名 */
    String value() default "";

    /** 字段是否允许为空 */
    boolean notNull() default false;
}