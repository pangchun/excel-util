package com.pangchun.poi.test;

import com.pangchun.poi.read.CommonRead;
import com.pangchun.poi.support.bean.ImageBean;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.List;
import java.util.Map;

/**
 * @author pangchun
 * @since 2021/6/6
 * @description 测试读取excel
 */
public class TestRead {

    /** excel文件名 */
    private static final String FILE_NAME = "F:\\码云\\excel-read-write-util\\excel-util\\src\\main\\resources\\static\\template\\person-template.xlsx";

    /** 图片上传位置 */
    private static final String IMAGE_PATH = "F:\\码云\\excel-read-write-util\\excel-util\\src\\main\\resources\\static\\image";


    public static void main(String[] args) throws IOException, InstantiationException, IllegalAccessException {
        File file = new File(FILE_NAME);
        // 定义表头行数
        int headRowNumber = 2;
        // 获取公共读取类
        CommonRead<Employee> commonRead = new CommonRead<>();
        // 获得工作簿
        Workbook workbook = commonRead.getWorkbook(file);
        // 获得工作表
        Sheet sheet = workbook.getSheetAt(0);
        // 获取表头内容，打印到控制台
        Map<Integer, Map<Integer, String>> headMap = commonRead.getHeadMap(sheet, headRowNumber);
        headMap.entrySet().forEach(integerMapEntry -> {
            System.out.println(integerMapEntry.getValue());
        });
        // 获取正文内容
        List<Employee> employeeList = commonRead.explainSheet(workbook, sheet, headRowNumber, Employee.class);
        employeeList.forEach(employee -> {
            System.out.println(employee.toString());
        });
        // 获取图片数据
        List<ImageBean> beanList = commonRead.explainPicture(workbook, sheet, IMAGE_PATH);
        beanList.forEach(imageBean -> {
            System.out.println(imageBean.toString());
        });
    }
}