package com.mfz.study.util.exceldesign;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.Map;

/**
 * @Description: 设置样式库-抽象
 * @Author mengfanzhu
 * @Date 2019/12/18 11:00
 * @Version 1.0
 */
public abstract class AbstractBuildStyle {

    /**
     * 任何实现类，如果定义的是内容样式，存入map的key必须是这个，否则不生效
     */
    public static final String CONTENT_STYLE_KEY = "contentStyleKey";

    /**
     * 任何实现类，如果定义的是标题样式且使用默认的设置Title方式，
     * 则存入map的key必须是这个，否则不生效，如果同时自定义了设置Title的实现，则不受这个限制
     */
    public static final String TITLE_STYLE_KEY = "titleStyleKey";

    /**
     * 任何实现类，如果定义的是表头样式，存入map的key必须是这个，否则不生效
     */
    public static final String HEAD_STYLE_KEY = "headStyleKey";

    /**
     * @param workbook 表格
     * @param styleMap 样式库
     */
    abstract void setStyle(SXSSFWorkbook workbook, Map<String,CellStyle> styleMap);
}
