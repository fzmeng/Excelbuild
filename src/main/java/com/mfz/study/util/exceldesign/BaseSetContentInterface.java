package com.mfz.study.util.exceldesign;

import java.util.List;
import java.util.Map;

/**
 * @Description: invokeData 需要实现此接口
 * @Author mengfanzhu
 * @Date 2019/12/20 10:11
 * @Version 1.0
 */
public interface BaseSetContentInterface {

    /**
     * @param json       数据构造入参
     * @param limitStart 数据分页开始
     * @param size       每页数据量
     * @return
     * @see DefaultBuildContentInvokeData
     */
    List<? extends Map<String, Object>> getExcelExportData(String json, Integer limitStart, Integer size);
}
