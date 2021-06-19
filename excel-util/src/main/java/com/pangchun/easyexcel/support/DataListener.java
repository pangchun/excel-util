package com.pangchun.easyexcel.support;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.fastjson.JSON;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author pangchun
 * @since 2021/6/19
 * @description 数据解析监听器
 */
public class DataListener<T> extends AnalysisEventListener<T> {

    private static final org.slf4j.Logger Log = LoggerFactory.getLogger(DataListener.class);

    /** 正文起始行(也表示有多少行表头) */
    private int headRowNumber;

    /** 对象数据 */
    private List<T> data = new ArrayList<T>();

    /** 合并单元格数据 */
    private List<CellExtra> mergedCellList = new ArrayList<CellExtra>();

    /** 构造函数 */
    public DataListener(int headRowNumber) {
        this.headRowNumber = headRowNumber;
    }

    /**
     * 每条记录解析时都会来调用此方法
     * @param object excel表中的一条记录对应一个java对象
     * @param analysisContext 上下文解析
     */
    @Override
    public void invoke(T object, AnalysisContext analysisContext) {
        data.add(object);
    }

    /**
     * 所有记录解析完成后会来调用此方法
     * @param analysisContext 上下文解析
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {}

    /**
     * 解析批注、超链接、合并单元格
     * @param extra 额外的单元格
     * @param context 上下文解析
     */
    @Override
    public void extra(CellExtra extra, AnalysisContext context) {
        switch (extra.getType()) {
            case COMMENT:
            case HYPERLINK:
                break;
            case MERGE:
                /* 这里要判断表头，如果是表头，则不存储 */
                if (extra.getRowIndex() >= headRowNumber) {
                    Log.info("解析到一个非表头合并单元格，覆盖区间为： ({},{})，({}，{})", extra.getFirstRowIndex(), extra.getFirstColumnIndex(), extra.getLastRowIndex(), extra.getLastColumnIndex() );
                    mergedCellList.add(extra);
                }
                break;
            default:
        }
    }

    /**
     * 解析表头数据
     * @param headMap 表头信息
     * @param context 上下文解析
     */
    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        Log.info("解析到一条表头数据： {}", JSON.toJSONString(headMap));
    }

    /**
     * 获取合并单元格数据
     * @return 合并单元格集合
     */
    public List<CellExtra> getMergedCellList() {
        return this.mergedCellList;
    }

    /**
     * 获取对象数据
     * @return 对象数据集合
     */
    public List<T> getAnalysisData() {
        return this.data;
    }
}