package com.pangchun.poi.support.annotation;

import java.lang.annotation.*;

/**
 * 自定义注解，用于映射excel表中字段位置，以便通过反射赋值。
 */
/**
 * @author pangchun
 * @since 2021/6/5
 * @description 行索引注解，表示实体属性与单元格的行位置对应关系
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelRow {

    /** 行索引，从0开始 */
    int index() default -1;

    /** 行名 */
    String value() default "";

    /** 字段是否允许为空 */
    boolean notNull() default false;
}