package com.mfz.study.util.design;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.util.Collection;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * @Description: 建造者模式-抽象建造者
 * @Author mengfanzhu
 * @Date 2019/12/18 11:24
 * @Version 1.0
 */
@Slf4j
public abstract class AbstractExcelBuilder {


    /**
     * 数据主体开始行下标
     */
    int rowStartIndex = 0;

    /**
     * excel名称
     */
    String excelFileName;

    /**
     * 工作对象
     */
    SXSSFWorkbook workbook;

    /**
     * 表格
     */
    SXSSFSheet sheet;

    /**
     * 样式库
     */
    Map<String, CellStyle> styleMap = new HashMap<>();

    /**
     * 属性-标题对照
     */
    LinkedHashMap<String, String> fieldTitleTable = new LinkedHashMap<>();

    /**
     * 获取属性值
     */
    AbstractGetFieldValue abstractGetFieldValue;

    /**
     * 设置样式
     */
    AbstractBuildStyle abstractBuildStyle;

    /**
     * 设置标题+表头
     */
    AbstractBuildTitle abstractBuildTitle;
    /**
     * 设置内容
     */
    AbstractBuildContent content;

    /**
     * 数据源所在类，需要实现接口
     *
     * @see BaseSetContentInterface
     */
    Class invokeClass;

    /**
     * 数据源所在方法，入参
     *
     * @see BaseSetContentInterface
     */
    String methodParamJson;
    Integer methodParamLimitStart;
    Integer methodParamLimitSize;

    void addFieldTitleTable(String fieldName, String titleName) {
        fieldTitleTable.put(fieldName, titleName);
    }

    /**
     * 建造工作对象
     *
     * @param wb
     * @return
     */
    abstract AbstractExcelBuilder buildWorkbook(SXSSFWorkbook wb);

    /**
     * 建造表格
     *
     * @param sheetName
     * @return
     */
    abstract AbstractExcelBuilder buildSheet(String sheetName);

    /**
     * @param data               数据源
     * @param heads              表头
     * @param columnTypeMap      列对应数据类型
     * @param abstractSetContent
     * @return
     * @see AbstractBuildContent
     */
    abstract AbstractExcelBuilder buildContent(Collection data, String[] heads, LinkedHashMap<String, Integer> columnTypeMap, Class<? extends AbstractBuildContent> abstractSetContent);

    /**
     * 构建标题+表头
     *
     * @param titles
     * @param heads
     * @return
     */
    abstract AbstractExcelBuilder buildTitlesAndHeads(String[] titles, String[] heads);

    /**
     * 文件名称
     *
     * @param excelFileName
     * @return
     */
    abstract AbstractExcelBuilder buildExcelFileName(String excelFileName);

    abstract AbstractExcelBuilder buildStyle();

    /**
     * @param clazz
     * @return
     * @see BaseSetContentInterface
     */
    abstract AbstractExcelBuilder buildInvokeClass(Class clazz);

    /**
     * @param methodParamJson
     * @param methodParamLimitStart
     * @param methodParamLimitSize
     * @return
     * @see BaseSetContentInterface
     */
    abstract AbstractExcelBuilder buildInvokeMethodParamValue(String methodParamJson, Integer methodParamLimitStart, Integer methodParamLimitSize);

    /**
     * 输出流
     *
     * @param response
     */
    abstract void flush(HttpServletResponse response);

    abstract void flush(HttpServletResponse response, String excelFileName);


    AbstractExcelBuilder buildColumn(int columnSize, boolean isAutoColumnWidth) {
        if (columnSize == 0) {
            return this;
        }
        if (isAutoColumnWidth) {
            for (int i = 0; i < columnSize; i++) {
                sheet.autoSizeColumn(i);
            }
            return this;
        }
        for (int i = 0; i < columnSize; i++) {
            sheet.setDefaultColumnWidth(i);
        }

        return this;
    }

    /**
     * 获取属性值的方式
     *
     * @param getFieldValue
     */
    void getFieldValueImpl(AbstractGetFieldValue getFieldValue) {
        if (null == getFieldValue) {
            this.abstractGetFieldValue = new DefaultGetValue();
        } else {
            this.abstractGetFieldValue = getFieldValue;
        }
    }

    void setStyleImpl(AbstractBuildStyle setStyle) {
        if (null == setStyle) {
            this.abstractBuildStyle = new DefaultBuildStyle();
        } else {
            this.abstractBuildStyle = setStyle;
        }
    }

    void setTitleImpl(AbstractBuildTitle setTitle) {
        if (null == setTitle) {
            this.abstractBuildTitle = new DefaultBuildTitle();
        } else {
            this.abstractBuildTitle = setTitle;
        }
    }

    void setContentImpl(Class<? extends AbstractBuildContent> abstractSetContent) {
        try {
            if (null == abstractSetContent) {
                this.content = new DefaultBuildContentModel(styleMap, sheet, abstractGetFieldValue).initParams(rowStartIndex);
                return;
            }
            AbstractBuildContent content = abstractSetContent.newInstance();
            if (content instanceof DefaultBuildContentInvokeData) {
                this.content = new DefaultBuildContentInvokeData(styleMap, sheet, abstractGetFieldValue)
                        .initParams(invokeClass, methodParamJson, methodParamLimitStart, methodParamLimitSize, rowStartIndex);
            } else if (content instanceof DefaultBuildContentMap) {
                this.content = new DefaultBuildContentMap(styleMap, sheet, abstractGetFieldValue).initParams(rowStartIndex);
            } else if (content instanceof DefaultBuildContentModel) {
                this.content = new DefaultBuildContentModel(styleMap, sheet, abstractGetFieldValue).initParams(rowStartIndex);
            } else {
                throw new IllegalArgumentException("Invalid AbstractBuildContent class: " + abstractSetContent.getName());
            }
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    }

    public SXSSFWorkbook getWorkbook() {
        return workbook;
    }
}

