package com.pangchun.poi.test;

import com.pangchun.poi.support.annotation.ExcelColumn;
import lombok.Data;
import lombok.experimental.Accessors;

import java.util.List;

/**
 * @author pangchun
 * @since 2021/6/6
 * @description 解析实体
 */
@Data
@Accessors(chain = true)
public class Employee implements Cloneable{

    @ExcelColumn(index = 0, value = "编号")
    private String id;

    @ExcelColumn(index = 1, value = "姓名")
    private String name;

    @ExcelColumn(index = 2, value = "性别")
    private String sex;

    @ExcelColumn(index = 3, value = "出生日期")
    private String birth;

    @ExcelColumn(index = 4, value = "通讯地址")
    private String address;

    @ExcelColumn(index = 5, value = "联系方式")
    private String phoneNumber;

    @ExcelColumn(index = 6, value = "所在部门")
    private String departmentName;

    @ExcelColumn(index = 7, value = "基础薪资")
    private String salary;

    @ExcelColumn(index = 8, value = "基础薪资抽取百分比")
    private String percent;

    @ExcelColumn(index = 9, value = "基础薪资(抽取后)")
    private String salaryAfterPercent;

    /** 在实体有多张图片时使用List<String> */
    @ExcelColumn(index = 10, value = "证件照")
    private String imageUrl;

    @Override
    public Employee clone() throws CloneNotSupportedException {
        return (Employee) super.clone();
    }
}