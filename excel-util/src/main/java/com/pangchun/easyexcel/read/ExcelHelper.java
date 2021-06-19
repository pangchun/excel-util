package com.pangchun.easyexcel.read;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.CellExtra;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.List;

/**
 * @author pangchun
 * @since 2021/6/19
 * @description excel解析辅助类
 */
public class ExcelHelper<T> {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelHelper.class);

    /**
     * 处理合并单元格
     * @param cellExtras 合并单元格集合
     * @param data 解析的数据集合
     * @param headRowNumber 表头行数
     * @return 填充好的数据
     */
    public List<T> resolveMergedCell(List<T> data, List<CellExtra> cellExtras, Integer headRowNumber) {
        for (CellExtra extra : cellExtras) {
            Integer firstRowIndex = extra.getFirstRowIndex() - headRowNumber;
            Integer lastRowIndex = extra.getLastRowIndex() - headRowNumber;
            Integer firstColumnIndex = extra.getFirstColumnIndex();
            Integer lastColumnIndex = extra.getLastColumnIndex();
            Object initValue = getInitValueFromList(firstRowIndex, firstColumnIndex, data);
            for (int i = firstRowIndex; i <= lastRowIndex; i++) {
                for (int j = firstColumnIndex; j <= lastColumnIndex; j++) {
                    setInitValueTOList(initValue, i, j, data);
                }
            }
        }
        return data;
    }

    /**
     * 设置合并单元格的值
     * @param initValue 初始值
     * @param rowIndex 行
     * @param columnIndex 列
     * @param data 解析数据
     */
    private void setInitValueTOList(Object initValue, Integer rowIndex, Integer columnIndex, List<T> data) {
        T item = data.get(rowIndex);
        for (Field field : item.getClass().getDeclaredFields()) {
            field.setAccessible(true);
            ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
            if (annotation != null) {
                if (annotation.index() == columnIndex) {
                    try {
                        field.set(item, initValue);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                        LOGGER.error("设置合并单元格数据初始值发生异常");
                    }
                }
            }
        }
    }

    /**
     * 获取合并单元格的初始值
     * 因为合并单元格的值是默认存在首行首列的单元格，并且firstRowIndex减去表头的行数后正好对应list<T>中的下标，列的位置又可以通过添加注解获得，这样就能拿到合并单元格的初始值了
     * @param firstRowIndex 合并单元格首行行数减去表头行数
     * @param firstColumnIndex 合并单元格首列
     * @param data easy-excel解析的list<T>数据
     * @return 合并单元格的初始值
     */
    private Object getInitValueFromList(Integer firstRowIndex, Integer firstColumnIndex, List<T> data) {
        Object initValue = null;
        T item = data.get(firstRowIndex);
        for(Field field : item.getClass().getDeclaredFields()) {
            field.setAccessible(true);
            ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
            if (annotation != null) {
                if (annotation.index() == firstColumnIndex) {
                    try {
                        initValue = field.get(item);
                        LOGGER.info("解析到合并单元格初始值{}", initValue);
                        break;
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                        LOGGER.error("解析合并单元格数据初始值发生异常");
                    }
                }
            }
        }
        return initValue;
    }

}