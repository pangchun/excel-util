package com.pangchun.easyexcel.test;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.enums.CellExtraTypeEnum;
import com.alibaba.excel.metadata.CellExtra;
import com.pangchun.easyexcel.read.ExcelHelper;
import com.pangchun.easyexcel.support.DataListener;

import java.util.List;

public class Main {

    private static final String path = "C:\\Users\\Administrator\\Desktop\\笔记\\excel-demo\\src\\main\\resources\\read.xlsx";

    public static void main(String[] args) {
        long start = System.currentTimeMillis();

        DataListener<Person> listener = new DataListener<>(4);

        EasyExcel.read(path, Person.class, listener)
                //读取批注、超链接、合并单元格需要在此处开启，默认是不读取的
                .extraRead(CellExtraTypeEnum.COMMENT)
                .extraRead(CellExtraTypeEnum.HYPERLINK)
                .extraRead(CellExtraTypeEnum.MERGE).sheet().headRowNumber(4).doRead();

        List<CellExtra> cellList = listener.getMergedCellList();

        List<Person> data = listener.getAnalysisData();

        ExcelHelper<Person> helper = new ExcelHelper<>();

        List<Person> personList = helper.resolveMergedCell(data, cellList, 4);

        for (Person person : personList) {
            System.out.println(person);
        }

        long end = System.currentTimeMillis();

        System.out.println("用时 " + (end - start));
    }
}