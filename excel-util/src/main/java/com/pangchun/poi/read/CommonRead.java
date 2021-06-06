package com.pangchun.poi.read;

import com.google.common.collect.Maps;
import com.pangchun.poi.support.bean.ImageBean;
import com.pangchun.poi.support.exception.ExcelReadException;
import com.pangchun.poi.support.annotation.ExcelColumn;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.joda.time.DateTime;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.util.*;

/**
 * @author pangchun
 * @since 2021/6/5
 * @description 通用的excel读取
 */
public class CommonRead<T> {

    /** 获取日志 */
    private static final Logger LOG = LoggerFactory.getLogger(CommonRead.class);

    /**
     * 获取工作簿
     * @param file excel文件
     * @return 工作簿
     */
    public Workbook getWorkbook(File file) throws IOException {
        Workbook workbook = null;
        String suffix03 = ".xls";
        String suffix07 = ".xlsx";
        String fileName = file.getName();
        if (fileName.endsWith(suffix03)) {
            workbook = new HSSFWorkbook(new FileInputStream(file));
        } else if (fileName.endsWith(suffix07)) {
            workbook = new XSSFWorkbook(new FileInputStream(file));
        }
        if (workbook != null) { workbook.close();}
        return workbook;
    }

    /**
     * 获取工作表
     * @param workbook 工作簿
     * @param sheetNo 表号
     * @return 工作表
     * @description 需要加入自己的逻辑判断就使用此方法修改，不需要加入其它逻辑时，直接使用getSheetAt(int var1)方法即可
     */
    public Sheet getSheet(Workbook workbook, int sheetNo) {
        Sheet sheet = null;
        int numberOfSheets = workbook.getNumberOfSheets();
        if (sheetNo <= numberOfSheets && sheetNo >= 0) { sheet = workbook.getSheetAt(sheetNo); }
        return sheet;
    }

    /**
     * 获取表头内容
     * @param sheet 工作表
     * @param headRowNumber 表头行数
     * @return 表头内容
     */
    public Map<Integer, Map<Integer, String>> getHeadMap(Sheet sheet, Integer headRowNumber) {
        // 参考了谷歌提供的guava包的指定集合初始值大小的创建方式Maps.newHashMapWithExpectedSize(int expectedSize)
        Map<Integer, Map<Integer, String>> headMap = new HashMap<>((int)((float)headRowNumber / 0.75F + 1.0F));
        for (int rowNo = 0; rowNo < headRowNumber; rowNo++) {
            Row row = sheet.getRow(rowNo);
            int numberOfCells = row.getPhysicalNumberOfCells();
            Map<Integer, String> dataMap = Maps.newHashMapWithExpectedSize(numberOfCells);
            for (int colNo = 0; colNo < numberOfCells; colNo++) {
                Cell cell = row.getCell(colNo);
                if (cell != null) {
                    String cellValue = cell.getStringCellValue();
                    // replace(" ", "").equals("")在jdk11以上可以使用isBlank()替换
                    if (cellValue != null && !"".equals(cellValue.replace(" ", ""))) {
                        dataMap.put(colNo, cellValue);
                    }
                }
            }
            if (!dataMap.isEmpty()) {
                headMap.put(rowNo, dataMap);
            }
        }
        return headMap;
    }

    /**
     * 获取表头合并单元格集合
     * @param sheet 工作表
     * @param headRowNumber 表头行数
     * @return 合并单元格集合
     */
    public List<CellRangeAddress> getHeadMerges(Sheet sheet, Integer headRowNumber) {
        // 当不能确定初始容量合适大小，设置默认16
        List<CellRangeAddress> headMerges = new ArrayList<>(16);
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress mergedRegion : mergedRegions) {
            int firstRow = mergedRegion.getFirstRow();
            if (firstRow < headRowNumber) {
                headMerges.add(mergedRegion);
            }
        }
        return headMerges;
    }

    /**
     * 获取非表头合并单元格集合
     * @param sheet 表
     * @param headRowNumber 表头行数
     * @return 合并单元格集合
     */
    public List<CellRangeAddress> getMerges(Sheet sheet, Integer headRowNumber) {
        List<CellRangeAddress> merges = new ArrayList<>(16);
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress mergedRegion : mergedRegions) {
            int firstRow = mergedRegion.getFirstRow();
            if (firstRow >= headRowNumber) {
                merges.add(mergedRegion);
            }
        }
        return merges;
    }

    /**
     * 获取单元格的值
     * @param workbook 工作簿
     * @param cell 单元格
     * @return 值
     */
    public Object getCellValue(Workbook workbook, Cell cell) {
        Object cellValue = null;
        if (cell != null) {
            CellType type = cell.getCellType();
            switch (type) {
                // 字符串（文本）单元格类型
                case STRING:
                    cellValue = cell.getStringCellValue();
                    break;
                // 布尔单元格类型
                case BOOLEAN:
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                // 公式单元格类型
                case FORMULA:
                    FormulaEvaluator evaluator = getFormulaEvaluator(workbook, cell);
                    cellValue = evaluator.evaluate(cell).formatAsString();
                    break;
                // 数字单元格类型（整数、小数、日期）
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        // 将日期转为指定的String格式
                        Date date = cell.getDateCellValue();
                        cellValue = new DateTime(date).toString("yyyy-MM-dd HH:mm:ss");
                    } else {
                        // 防止数字过长，转为String，保留两位小数，若小数为零则直接取整数
                        DecimalFormat format = null;
                        double value = cell.getNumericCellValue();
                        format = new DecimalFormat("0.00");
                        String valueString = format.format(value);
                        String replace = valueString.substring(valueString.lastIndexOf(".") + 1).replace("0", "");
                        if ("".equals(replace)) {
                            format = new DecimalFormat("0");
                            cellValue = format.format(value);
                        } else {
                            cellValue = valueString;
                        }
                    }
                    break;
                // 空白单元格类型
                case BLANK:
                // 错误单元格类型
                case ERROR:
                // 未知类型，用于表示初始化之前的状态或缺少具体类型。 仅限内部使用。
                case _NONE:
                default: break;
            }
        }
        return cellValue;
    }

    /**
     * 获取计算公式
     * @param workbook 工作簿
     * @param cell 单元格
     * @return 公式
     */
    public FormulaEvaluator getFormulaEvaluator(Workbook workbook, Cell cell) {
        FormulaEvaluator evaluator = null;
        if (workbook instanceof HSSFWorkbook) {
            evaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);
        } else if (workbook instanceof XSSFWorkbook) {
            evaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
        } else {
            throw new ExcelReadException("不支持其他类型工作簿，需要请自行添加");
        }
        return evaluator;
    }

    /**
     * 获取表头以下的正文数据，会将一行记录解析为一个T对象，注意这里这能解析文字，图片是单独设置到java对象中的
     * @param workbook 工作簿
     * @param sheet 工作表
     * @param headRowNumber 表头行数
     * @param clazz 要解析成的Java对象类型，传参如 `Person.class`
     * @return 用java对象形式表示的对象集合
     */
    public List<T> explainSheet(Workbook workbook, Sheet sheet, Integer headRowNumber, Class<T> clazz) throws IllegalAccessException, InstantiationException {
        int numberOfRows = sheet.getPhysicalNumberOfRows();
        List<T> data = new ArrayList<>((int)((float)numberOfRows / 0.75F + 1.0F));
        // 设置字段的值
        for (int rowNo = headRowNumber; rowNo < numberOfRows; rowNo++) {
            Row row = sheet.getRow(rowNo);
            if (row != null) {
                T instance = null;
                instance = clazz.newInstance();
                int numberOfCells = row.getPhysicalNumberOfCells();
                for (int colNo = 0; colNo < numberOfCells; colNo++) {
                    Object cellValue = getCellValue(workbook, row.getCell(colNo));
                    setFieldValue(instance, colNo, cellValue);
                }
                data.add(instance);
            }
        }
        // 设置合并单元格的值
        List<CellRangeAddress> addressList = getMerges(sheet, headRowNumber);
        for (CellRangeAddress address : addressList) {
            int firstRowIndex = address.getFirstRow() - headRowNumber;
            int lastRowIndex = address.getLastRow() - headRowNumber;
            int firstColumnIndex = address.getFirstColumn();
            int lastColumnIndex = address.getLastColumn();
            Object initValue = getInitValueOfMergeFromList(firstRowIndex, firstColumnIndex, data);
            for (int rowNo = firstRowIndex; rowNo <= lastRowIndex; rowNo++) {
                setInitValueOfMergeToList(initValue, rowNo, firstColumnIndex, data);
                 // 这里是处理列合并的数据
                for (int colNo = firstColumnIndex; colNo <= lastColumnIndex; colNo++) {
                    setInitValueOfMergeToList(initValue, rowNo, colNo, data);
                }
            }
        }
        return data;
    }

    /**
     * 设置实体字段的值，通过反射遍历设置，将excel解析到的每个值设置到实体的属性中
     * @param instance 要设置值的实体对象
     * @param columnIndex cellValue对应excel表中的列索引
     * @param cellValue excel表格的值
     */
    public void setFieldValue(Object instance, Integer columnIndex, Object cellValue) throws IllegalAccessException {
        for (Field field : instance.getClass().getDeclaredFields()) {
            field.setAccessible(true);
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            if (annotation != null) {
                if (annotation.index() == columnIndex) {
                    field.set(instance, cellValue);
                }
            }
        }
    }

    /**
     * 获取合并单元格的初始值
     * 因为合并单元格的值是默认存在首行首列的单元格，并且firstRowIndex减去表头的行数后正好对应list<T>中的下标，列的位置又可以通过添加注解获得，这样就能拿到合并单元格的初始值了
     * @param firstRowIndex 合并单元格首行行数减去表头行数
     * @param firstColumnIndex 合并单元格首列
     * @param data explainSheet方法解返回的list<T>数据
     * @return 合并单元格的初始值
     */
    private Object getInitValueOfMergeFromList(Integer firstRowIndex, Integer firstColumnIndex, List<T> data) throws IllegalAccessException {
        Object initValue = null;
        T item = data.get(firstRowIndex);
        for(Field field : item.getClass().getDeclaredFields()) {
            field.setAccessible(true);
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            if (annotation != null) {
                if (annotation.index() == firstColumnIndex) {
                    initValue = field.get(item);
                }
            }
        }
        return initValue;
    }

    /**
     * 设置合并单元格的值到List<T>中
     * @param initValue 初始值
     * @param rowNo 行号
     * @param colNo 列号
     * @param data 解析数据
     */
    private void setInitValueOfMergeToList(Object initValue, Integer rowNo, Integer colNo, List<T> data) throws IllegalAccessException {
        T item = data.get(rowNo);
        for (Field field : item.getClass().getDeclaredFields()) {
            field.setAccessible(true);
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            if (annotation != null) {
                if (annotation.index() == colNo) {
                    field.set(item, initValue);
                }
            }
        }
    }

    /**
     * 解析excel表中的图片数据
     * @param workbook 工工作簿
     * @param sheet 工作表
     * @param path 图片上传路径
     * @return 返回ImageBean集合
     */
    public List<ImageBean> explainPicture(Workbook workbook, Sheet sheet, String path) throws IOException {
        List<ImageBean> imageBeans = new ArrayList<>();
        if (workbook instanceof HSSFWorkbook) {
            List<HSSFShape> shapes = ((HSSFSheet) sheet).getDrawingPatriarch().getChildren();
            for (HSSFShape shape : shapes) {
                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                if (shape instanceof HSSFPicture) {
                    // 图片位置
                    HSSFPicture picture = (HSSFPicture) shape;
                    // 获取图片左上角的行号
                    int firstRowIndex = anchor.getRow1();
                    // 获取图片右下角的行号
                    int lastRowIndex = anchor.getRow2();
                    // 获取图片左上角的列号
                    int firstColIndex = anchor.getCol1();
                    // 获取图片右下角的列号
                    int lastColIndex = anchor.getCol2();
                    // 上传图片
                    String url = uploadImage(picture, path);
                    // 将图片信息封装到ImageBean中
                    ImageBean imageBean = getImageBean(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex, url, sheet);
                    imageBeans.add(imageBean);
                }
            }
        } else if (workbook instanceof XSSFWorkbook) {
            List<XSSFShape> shapes = ((XSSFSheet) sheet).getDrawingPatriarch().getShapes();
            for (XSSFShape shape : shapes) {
                XSSFClientAnchor anchor = (XSSFClientAnchor) shape.getAnchor();
                if (shape instanceof XSSFPicture) {
                    XSSFPicture picture = (XSSFPicture) shape;
                    int firstRowIndex = anchor.getRow1();
                    int lastRowIndex = anchor.getRow2();
                    int firstColIndex = anchor.getCol1();
                    int lastColIndex = anchor.getCol2();
                    String url = uploadImage(picture, path);
                    ImageBean imageBean = getImageBean(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex, url, sheet);
                    imageBeans.add(imageBean);
                }
            }
        }
        return imageBeans;
    }

    /**
     * 上传单个图片并返回访问路径， 这里的文件上传是模拟上传到本地的，真实工作一般是调用公共服务上传到图片服务器，根据实际情况修改即可。
     * @param picture 图片
     * @param path 上传路径
     * @return 访问路径
     */
    public String uploadImage(Picture picture, String path) throws IOException {
        PictureData pictureData = picture.getPictureData();
        byte[] data = pictureData.getData();
        String ext = pictureData.suggestFileExtension();
        path = path + File.separator + Math.random() + "." +ext;
        File file = new File(path);
        // try-with-resource写法能自动关闭流
        try (FileOutputStream stream = new FileOutputStream(file)) {
            stream.write(data);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return path;
    }

    /**
     * 将图片信息封装到ImageBean中，之后根据ImageBean中的行列属性将图片设置到对应的List<T>中即可
     * 注意：需要判断是否为合并单元格, 为合并单元格时，对应行列属性应设置合并单元格的行列号
     * @param firstRowIndex 图片所在单元格开始行
     * @param lastRowIndex 图片所在单元格结束行
     * @param firstColIndex 图片所在单元格开始列
     * @param lastColIndex 图片所在单元格结束列
     * @param sheet 工作表
     * @return 封装了图片信息的ImageBean
     */
    public ImageBean getImageBean(int firstRowIndex, int lastRowIndex, int firstColIndex, int lastColIndex, String url, Sheet sheet) {
        ImageBean imageBean = new ImageBean();
        // 先默认此图片不在任何一个单元格内
        boolean matchNoCell = true;
        // 判断图片是否在合并单元格内，若在合并单元格内, 图片的边界值取其所在合并单元格的边界值
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress mergedRegion : mergedRegions) {
            int firstRow = mergedRegion.getFirstRow();
            int lastRow = mergedRegion.getLastRow();
            int firstColumn = mergedRegion.getFirstColumn();
            int lastColumn = mergedRegion.getLastColumn();
            if (firstRowIndex >= firstRow
                    && firstColIndex >= firstColumn
                    && lastRowIndex <= lastRow
                    && lastColIndex <= lastColumn) {
                // 图片所在单元格在此合并单元格内
                firstRowIndex = firstRow;
                lastRowIndex = lastRow;
                firstColIndex = firstColumn;
                lastColIndex = lastColumn;
                matchNoCell = false;
                break;
            }
        }
        // 不在合并单元格中，判断是否在非合并单元格内
        if (matchNoCell) {
            if (!(firstRowIndex == lastRowIndex && firstColIndex == lastColIndex)){
                // TODO 图片解析中如果抛出异常，已经上传的图片无法回滚，只能去手动删除，因此需要提供一个删除接口
                throw new ExcelReadException("图片的插入位置有误，图片应位于所在单元格的边界线内，不应跨越两个单元格。");
            }
        }
        imageBean.setFirstRow(firstRowIndex).setLastRow(lastRowIndex).setFirstCol(firstColIndex).setLastCol(lastColIndex).setUrl(url);
        return imageBean;
    }

}
