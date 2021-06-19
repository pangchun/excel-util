package com.pangchun.easyexcel.test;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.util.Date;

@Data
public class Person {

    @ExcelProperty(index = 0)
    private String name;
    @ExcelProperty(index = 1)
    private Date birth;
    @ExcelProperty(index = 2)
    private Long idCardNumber;
    @ExcelProperty(index = 3)
    private String hobby;

    @Override
    public String toString() {
        return "Person{" +
                "name='" + name + '\'' +
                ", birth=" + birth +
                ", hobby='" + hobby + '\'' +
                ", idCardNumber=" + idCardNumber +
                '}';
    }
}