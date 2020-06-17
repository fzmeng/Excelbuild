package com.mfz.study.util.exceldesign;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.util.Assert;

import javax.servlet.http.HttpServletResponse;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.Set;

/**
 * @Description: 建造者模式-负责规范流程
 * @Author mengfanzhu
 * @Date 2019/12/18 13:59
 * @Version 1.0
 */
public class ExcelDirector {

    private AbstractExcelBuilder abstractExcelBuilder = new ExcelBuilder();

    private String sheetName;

    private String excelFileName;

    private String[] heads;

    private String[] titles;

    private LinkedHashMap<String, Integer> columnTypeMap;

    private Collection data;

    private SXSSFWorkbook workbook;

    private boolean isAutoColumnWidth = false;

    private Integer columnSize = null;

    /**
     * 数据源所在类，方法入参
     * implements
     *
     * @see BaseSetContentInterface
     */
    private Class clazz;
    private String methodParamJson;
    private Integer methodParamLimitStart;
    private Integer methodParamLimitSize;

    public ExcelDirector() {
        overrideDefaultGetFieldValue(null)
                .overrideDefaultSetTitle(null)
                .overrideDefaultSetStyle(null);
    }

    /**
     * 构建一个对象的流程
     *
     * @return
     * @see TestController
     */
    public ExcelDirector constructModel() {
        Assert.notNull(heads, "constructModel -> heads need to do ");
        Assert.notNull(data, "constructModel -> data need to do ");
        abstractExcelBuilder
                .buildWorkbook(workbook)
                .buildSheet(sheetName)
                .buildColumn(columnSize, isAutoColumnWidth)
                .buildStyle()
                .buildTitlesAndHeads(titles, heads)
                .buildExcelFileName(excelFileName)
                .buildContent(data, heads, null, DefaultBuildContentModel.class);
        return this;
    }

    /**
     * //指定列对应类型
     * LinkedHashMap<String, Integer> columnTypeMap = new LinkedHashMap<>();
     * columnTypeMap.put("id", 1); //1 -> String
     * columnTypeMap.put("name", 1);
     * columnTypeMap.put("address", 1);
     * //指定表头
     * LinkedHashMap<String, String> titleMap = new LinkedHashMap<>();
     * titleMap.put("id", "唯一标识符");
     * titleMap.put("name", "名字");
     * titleMap.put("address", "地址");
     * <p>
     * new ExcelDirector()
     * .addFieldTitleTable(titles)
     * //可不指定columnTypeMap 默认为String
     * .setColumnTypeMap(columnTypeMap)
     * .setData(result)
     * .setExcelFileName(filename)
     * .constructMap()
     * .flush(response);
     * <p>
     * 构建一个Map数据源的流程
     *
     * @return
     * @see TestController
     */
    public ExcelDirector constructMap() {
        if (columnTypeMap == null || columnTypeMap.size() != columnSize) {
            columnTypeMap = new LinkedHashMap<>();
            for (String fieldName : heads) {
                columnTypeMap.put(fieldName, CellType.STRING.getCode());
            }
        }
        Assert.notNull(heads, "constructMap -> heads need to do ");
        Assert.notNull(columnTypeMap, "constructMap -> columnTypeMap need to do ");
        Assert.notNull(data, "constructMap -> data need to do ");
        abstractExcelBuilder
                .buildWorkbook(workbook)
                .buildSheet(sheetName)
                .buildColumn(columnSize, isAutoColumnWidth)
                .buildStyle()
                .buildTitlesAndHeads(titles, heads)
                .buildExcelFileName(excelFileName)
                .buildContent(data, null, columnTypeMap, DefaultBuildContentMap.class);
        return this;
    }


    /**
     * @return
     * @see TestController
     * 构建一个Map 类型，反射数据源的流程
     * 示例：
     * 1.
     * //指定列对应类型
     * LinkedHashMap<String, Integer> columnTypeMap = new LinkedHashMap<>();
     * columnTypeMap.put("id", 1); //1 -> String
     * columnTypeMap.put("name", 1);
     * columnTypeMap.put("address", 1);
     * //指定表头
     * LinkedHashMap<String, String> titleMap = new LinkedHashMap<>();
     * titleMap.put("id", "唯一标识符");
     * titleMap.put("name", "名字");
     * titleMap.put("address", "地址");
     * <p>
     * SXSSFWorkbook wb = new ExcelDirector()
     * .addFieldTitleTable(titleMap)
     * .setColumnTypeMap(columnTypeMap)
     * .setBoundInvokeClass(this.getClass())
     * .setBoundInvokeMethodParamValues(JSONObject.toJSONString(params), 0, 5000)
     * .constructMapInvokeData()
     * .getWorkbook();
     * <p>
     * 2. 直接导出
     * new ExcelDirector()
     * .addFieldTitleTable(titleMap)
     * .setTitles("sheet1")
     * .setColumnTypeMap(columnTypeMap)
     * .setBoundInvokeClass(this.getClass())
     * .setBoundInvokeMethodParamValues(JSONObject.toJSONString(params),0,1000)
     * .constructMapInvokeData()
     * .flush(response,"fileName");
     */
    public ExcelDirector constructMapInvokeData() {
        Assert.notNull(clazz, "buildInvokeClass(clazz) need to do ");
        Assert.notNull(methodParamJson, "buildInvokeMethodParamValue(methodParamJson , methodParamLimitStart,methodParamLimitSize) need to do");
        Assert.notNull(methodParamLimitStart, "buildInvokeMethodParamValue(methodParamJson , methodParamLimitStart,methodParamLimitSize) need to do");
        Assert.notNull(methodParamLimitSize, "buildInvokeMethodParamValue(methodParamJson , methodParamLimitStart,methodParamLimitSize) need to do");

        abstractExcelBuilder
                .buildWorkbook(workbook)
                .buildSheet(sheetName)
                .buildColumn(columnSize, isAutoColumnWidth)
                .buildStyle()
                .buildTitlesAndHeads(titles, heads)
                .buildInvokeClass(clazz)
                .buildInvokeMethodParamValue(methodParamJson, methodParamLimitStart, methodParamLimitSize)
                .buildContent(null, null, columnTypeMap, DefaultBuildContentInvokeData.class);
        return this;
    }

    /**
     * 添加属性标题对照
     *
     * @param fieldTitle
     * @return
     */
    public ExcelDirector addFieldTitleTable(LinkedHashMap<String, String> fieldTitle) {
        if (fieldTitle == null || fieldTitle.size() == 0) {
            return null;
        }
        Set<String> strings = fieldTitle.keySet();
        for (String key : strings) {
            abstractExcelBuilder.addFieldTitleTable(key, fieldTitle.get(key));
        }
        setHeads(strings.toArray(new String[strings.size()]));
        return this;
    }

    public ExcelDirector addFieldTitleTable(String fieldName, String titleName) {
        abstractExcelBuilder.addFieldTitleTable(fieldName, titleName);
        return this;
    }


    public SXSSFWorkbook getWorkbook() {
        return abstractExcelBuilder.getWorkbook();
    }

    public ExcelDirector setWorkbook(SXSSFWorkbook workbook) {
        abstractExcelBuilder.workbook = workbook;
        return this;
    }

    /**
     * 覆盖-获取属性值
     *
     * @param getFieldValue
     * @return
     */
    public ExcelDirector overrideDefaultGetFieldValue(AbstractGetFieldValue getFieldValue) {
        abstractExcelBuilder.getFieldValueImpl(getFieldValue);
        return this;
    }

    public ExcelDirector isAutoColumnWidth(boolean isAutoColumnWidth) {
        this.isAutoColumnWidth = isAutoColumnWidth;
        return this;
    }

    /**
     * 覆盖-默认设置的样式
     *
     * @param setStyle
     * @return
     */
    public ExcelDirector overrideDefaultSetStyle(AbstractBuildStyle setStyle) {
        abstractExcelBuilder.setStyleImpl(setStyle);
        return this;
    }

    /**
     * 覆盖-默认设置标题
     *
     * @param setTitle
     * @return
     */
    public ExcelDirector overrideDefaultSetTitle(AbstractBuildTitle setTitle) {
        abstractExcelBuilder.setTitleImpl(setTitle);
        return this;
    }


    /**
     * 写到响应流
     *
     * @param response
     */
    public void flush(HttpServletResponse response) {
        abstractExcelBuilder.flush(response);
    }

    public void flush(HttpServletResponse response, String excelFileName) {
        abstractExcelBuilder.flush(response, excelFileName);
    }

    public ExcelDirector setSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public ExcelDirector setHeads(String[] heads) {
        if (columnSize == null) {
            columnSize = heads.length;
        }
        this.heads = heads;
        return this;
    }

    public Collection getData() {
        return data;
    }

    public ExcelDirector setData(Collection data) {
        this.data = data;
        return this;
    }

    /**
     * 设置列对应数据类型
     *
     * @param columnTypeMap
     * @return
     */
    public ExcelDirector setColumnTypeMap(LinkedHashMap<String, Integer> columnTypeMap) {
        this.columnTypeMap = columnTypeMap;
        return this;
    }

    /**
     * 设置标题
     *
     * @param titles
     * @return
     */
    public ExcelDirector setTitles(String... titles) {
        this.titles = titles;
        return this;
    }

    /**
     * 设置文件名称
     *
     * @param excelFileName
     * @return
     */
    public ExcelDirector setExcelFileName(String excelFileName) {
        this.excelFileName = excelFileName;
        return this;
    }


    /**
     * 设置反射类名
     *
     * @param invokeClass
     * @return
     */
    public ExcelDirector setBoundInvokeClass(Class invokeClass) {
        this.clazz = invokeClass;
        return this;
    }

    /**
     * 设置反射方法中参数值
     *
     * @param methodParamJson
     * @param methodParamLimitStart
     * @param methodParamLimitSize
     * @return
     * @see BaseSetContentInterface
     */
    public ExcelDirector setBoundInvokeMethodParamValues(String methodParamJson, Integer methodParamLimitStart, Integer methodParamLimitSize) {
        this.methodParamJson = methodParamJson;
        this.methodParamLimitStart = methodParamLimitStart;
        this.methodParamLimitSize = methodParamLimitSize;
        return this;
    }

}
